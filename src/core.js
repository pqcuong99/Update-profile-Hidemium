// Logic thuần Node (không phụ thuộc Electron): đọc Excel, chuyển đổi proxy, gọi API Hidemium.
const XLSX = require('xlsx');

const UUID_LIKE = /[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/i;
const PROXY_TYPES = ['HTTP', 'HTTPS', 'SOCKS4', 'SOCKS5', 'SSH'];

function cellText(value) {
  if (value === null || value === undefined) return '';
  return String(value);
}

// Ô trống hoặc chữ "null" => không gọi API cho cột đó.
function isSkipValue(value) {
  const t = cellText(value).trim();
  return t === '' || t.toLowerCase() === 'null';
}

function isValidPort(port) {
  if (!/^\d+$/.test(port)) return false;
  const n = Number(port);
  return n >= 1 && n <= 65535;
}

function buildProxy(type, host, port, user, pass) {
  type = (type || 'HTTP').toUpperCase();
  if (type === 'HTTPS') type = 'HTTP';
  if (!PROXY_TYPES.includes(type)) throw new Error(`Loại proxy không hỗ trợ: ${type}`);
  if (!host) throw new Error('Thiếu IP/host');
  if (!isValidPort(port)) throw new Error(`Port không hợp lệ: ${port}`);
  // Giống cách Hidemium tự tạo chuỗi proxy: bỏ các phần trống (không auth => HTTP|ip|port).
  return [type, host, port, user, pass].filter((p) => !!p).join('|');
}

/**
 * Chuyển proxy trong Excel sang định dạng Hidemium "TYPE|host|port|user|pass".
 * Hỗ trợ:
 *   127.0.0.1:60007
 *   103.182.23.51:60009:user:pass
 *   user:pass@103.182.23.51:60009
 *   http://user:pass@103.182.23.51:60009   (http, https, socks4, socks5, ssh)
 *   HTTP|1.1.1.1|1234|user|123              (giữ nguyên)
 */
function convertProxy(raw) {
  const s = cellText(raw).trim();

  if (s.includes('|')) {
    const [type, host, port, user, pass] = s.split('|').map((p) => p.trim());
    return buildProxy(type, host, port, user, pass);
  }

  const url = s.match(/^([a-z0-9]+):\/\/(?:([^:@/]*)(?::([^@/]*))?@)?([^:@/]+):(\d+)\/?$/i);
  if (url) {
    const [, scheme, user, pass, host, port] = url;
    return buildProxy(scheme, host, port, safeDecode(user), safeDecode(pass));
  }

  const at = s.match(/^([^:@]+):([^@]*)@([^:@]+):(\d+)$/);
  if (at) {
    const [, user, pass, host, port] = at;
    return buildProxy('HTTP', host, port, user, pass);
  }

  const parts = s.split(':');
  if (parts.length === 2) return buildProxy('HTTP', parts[0], parts[1]);
  if (parts.length === 4) return buildProxy('HTTP', parts[0], parts[1], parts[2], parts[3]);

  throw new Error(`Không nhận dạng được định dạng proxy: ${s}`);
}

function safeDecode(v) {
  if (!v) return v;
  try {
    return decodeURIComponent(v);
  } catch {
    return v;
  }
}

// Kế hoạch cập nhật cho 1 dòng: cột nào có giá trị thì mới gọi API.
function planRow(row) {
  const plan = { name: null, note: null, proxy: null };
  if (!isSkipValue(row.name)) plan.name = { value: cellText(row.name).trim() };
  if (!isSkipValue(row.note)) plan.note = { value: cellText(row.note) };
  if (!isSkipValue(row.proxy)) {
    try {
      plan.proxy = { value: convertProxy(row.proxy) };
    } catch (err) {
      plan.proxy = { error: err.message };
    }
  }
  return plan;
}

function readSheetRows(sheet) {
  const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null, raw: false, blankrows: true });
  const range = sheet['!ref'] ? XLSX.utils.decode_range(sheet['!ref']) : { s: { r: 0 } };
  const rows = [];
  matrix.forEach((r, i) => {
    const uid = cellText(r[0]).trim();
    if (!uid) return;
    // Dòng đầu là tiêu đề nếu cột A không phải UUID (VD: "UID").
    if (i === 0 && !UUID_LIKE.test(uid)) return;
    const row = {
      excelRow: range.s.r + i + 1,
      uid,
      name: cellText(r[1]),
      note: cellText(r[2]),
      proxy: cellText(r[3]),
      status: cellText(r[4]),
    };
    row.plan = planRow(row);
    rows.push(row);
  });
  return rows;
}

function readExcel(filePath) {
  const wb = XLSX.readFile(filePath, { cellDates: false });
  return {
    filePath,
    sheets: wb.SheetNames.map((name) => ({ name, rows: readSheetRows(wb.Sheets[name]) })),
  };
}

// Ghi trạng thái vào cột E (Status) của sheet và lưu ra file mới.
function exportResults({ sourcePath, sheetName, statuses, savePath }) {
  const wb = XLSX.readFile(sourcePath);
  const sheet = wb.Sheets[sheetName];
  if (!sheet) throw new Error(`Không tìm thấy sheet ${sheetName}`);
  const statusCol = 4; // cột E
  const headerAddr = XLSX.utils.encode_cell({ r: 0, c: statusCol });
  if (!sheet[headerAddr]) sheet[headerAddr] = { t: 's', v: 'Status' };
  let maxRow = 0;
  for (const [excelRow, text] of Object.entries(statuses)) {
    const r = Number(excelRow) - 1;
    sheet[XLSX.utils.encode_cell({ r, c: statusCol })] = { t: 's', v: text };
    maxRow = Math.max(maxRow, r);
  }
  const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1');
  range.e.c = Math.max(range.e.c, statusCol);
  range.e.r = Math.max(range.e.r, maxRow);
  sheet['!ref'] = XLSX.utils.encode_range(range);
  XLSX.writeFile(wb, savePath);
  return savePath;
}

// ---------------- API ----------------

function normalizeBaseUrl(baseUrl) {
  return String(baseUrl || 'http://127.0.0.1:2222').trim().replace(/\/+$/, '');
}

async function callApi({ baseUrl, isLocal, method, path, body, timeoutMs = 20000 }) {
  const url = `${normalizeBaseUrl(baseUrl)}${path}?is_local=${isLocal ? 'true' : 'false'}`;
  const started = Date.now();
  try {
    const res = await fetch(url, {
      method,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
      signal: AbortSignal.timeout(timeoutMs),
    });
    const text = await res.text();
    let json = null;
    try {
      json = JSON.parse(text);
    } catch {
      /* không phải JSON */
    }
    const payload = json && typeof json === 'object' && json.data && json.data.type ? json.data : json;
    const apiType = payload && typeof payload === 'object' ? payload.type || payload.status : undefined;
    const ok = res.ok && (!apiType || ['success', 'successfully'].includes(String(apiType).toLowerCase()));
    const message =
      (payload && (payload.title || payload.message || payload.msg)) || text.slice(0, 300) || `HTTP ${res.status}`;
    return { ok, httpStatus: res.status, message: String(message), url, method, body, ms: Date.now() - started };
  } catch (err) {
    const message =
      err.name === 'TimeoutError'
        ? `Hết thời gian chờ (${timeoutMs / 1000}s)`
        : err.cause && err.cause.code === 'ECONNREFUSED'
          ? 'Không kết nối được API (Hidemium đã mở chưa?)'
          : err.message;
    return { ok: false, httpStatus: 0, message, url, method, body, ms: Date.now() - started };
  }
}

const ACTIONS = {
  name: (uid, value) => ({
    method: 'PUT',
    path: '/v2/browser/update-once',
    body: { column: 'name', profile_uuid: uid, data: value },
  }),
  note: (uid, value) => ({
    method: 'PUT',
    path: '/v2/browser/update-note',
    body: { note: value, profile_uuid: uid },
  }),
  proxy: (uid, value) => ({
    method: 'POST',
    path: '/v2/browser/proxy/update',
    body: { browser_update: [{ uuid: uid, proxy: value }] },
  }),
};

async function updateRow(row, { baseUrl, isLocal, timeoutMs }) {
  const plan = row.plan || planRow(row);
  const results = {};
  for (const key of ['name', 'note', 'proxy']) {
    const item = plan[key];
    if (!item) {
      results[key] = { state: 'skip' };
      continue;
    }
    if (item.error) {
      results[key] = { state: 'error', message: item.error };
      continue;
    }
    const req = ACTIONS[key](row.uid, item.value);
    const res = await callApi({ baseUrl, isLocal, timeoutMs, ...req });
    results[key] = { state: res.ok ? 'ok' : 'error', message: res.message, request: req, httpStatus: res.httpStatus };
  }
  return results;
}

function summarizeResults(results) {
  const labels = { name: 'Name', note: 'Note', proxy: 'Proxy' };
  const parts = [];
  for (const key of ['name', 'note', 'proxy']) {
    const r = results[key];
    if (!r || r.state === 'skip') continue;
    parts.push(r.state === 'ok' ? `${labels[key]}: OK` : `${labels[key]}: Lỗi - ${r.message}`);
  }
  return parts.length ? parts.join(' | ') : 'Bỏ qua (không có dữ liệu)';
}

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

/**
 * Chạy tuần tự từng dòng. onProgress({ index, row, results, summary }) được gọi sau mỗi dòng.
 * shouldStop() trả true để dừng giữa chừng.
 */
async function runUpdates(rows, { baseUrl, isLocal, delayMs = 0, timeoutMs, onProgress, shouldStop }) {
  let done = 0;
  for (let i = 0; i < rows.length; i++) {
    if (shouldStop && shouldStop()) return { stopped: true, done };
    const row = rows[i];
    const results = await updateRow(row, { baseUrl, isLocal, timeoutMs });
    done++;
    if (onProgress) onProgress({ index: i, row, results, summary: summarizeResults(results) });
    if (delayMs > 0 && i < rows.length - 1) await sleep(delayMs);
  }
  return { stopped: false, done };
}

module.exports = {
  isSkipValue,
  convertProxy,
  planRow,
  readExcel,
  exportResults,
  callApi,
  updateRow,
  runUpdates,
  summarizeResults,
};
