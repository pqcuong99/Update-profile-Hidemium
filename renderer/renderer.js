'use strict';

const $ = (id) => document.getElementById(id);
const KEYS = ['name', 'note', 'proxy'];
const LABELS = { name: 'Name', note: 'Note', proxy: 'Proxy' };
const SETTINGS_KEY = 'excel-api-updater:settings';

const state = {
  file: null, // { filePath, sheets: [{ name, rows }] }
  sheetIndex: 0,
  results: new Map(), // id -> { results, summary }
  unchecked: new Set(), // id các dòng bỏ chọn
  running: false,
  runToken: null,
  runningId: null,
  queue: [],
  done: 0,
  progressBound: false,
};

// window.api do preload.js cung cấp; truy cập lúc dùng để không lỗi khi trang tải.
const getApi = () => {
  if (!window.api) throw new Error('Không tìm thấy API của ứng dụng (preload).');
  return window.api;
};

// ---------- tiện ích ----------

function esc(value) {
  return String(value ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);
}

function rowId(sheetName, row) {
  return `${sheetName}::${row.excelRow}`;
}

function currentSheet() {
  return state.file ? state.file.sheets[state.sheetIndex] : null;
}

function currentRows() {
  const sheet = currentSheet();
  return sheet ? sheet.rows : [];
}

function getMode() {
  return document.querySelector('input[name="mode"]:checked').value;
}

function fileNameOf(p) {
  return String(p).split(/[\\/]/).pop();
}

function loadSettings() {
  try {
    const s = JSON.parse(localStorage.getItem(SETTINGS_KEY) || '{}');
    if (s.mode) {
      const radio = document.querySelector(`input[name="mode"][value="${s.mode}"]`);
      if (radio) radio.checked = true;
    }
    if (s.baseUrl) $('baseUrl').value = s.baseUrl;
    if (s.delayMs !== undefined) $('delayMs').value = s.delayMs;
    return s;
  } catch {
    return {};
  }
}

function saveSettings() {
  try {
    const sheet = currentSheet();
    localStorage.setItem(
      SETTINGS_KEY,
      JSON.stringify({
        mode: getMode(),
        baseUrl: $('baseUrl').value,
        delayMs: $('delayMs').value,
        // File + sheet đang mở, để lần sau mở app tự tải lại.
        lastFile: state.file ? state.file.filePath : undefined,
        lastSheet: sheet ? sheet.name : undefined,
      }),
    );
  } catch {
    /* bỏ qua */
  }
}

function log(text, cls = '') {
  const el = $('log');
  const line = document.createElement('div');
  if (cls) line.className = cls;
  line.textContent = `[${new Date().toLocaleTimeString()}] ${text}`;
  el.appendChild(line);
  while (el.childElementCount > 1000) el.firstChild.remove();
  el.scrollTop = el.scrollHeight;
  $('logCount').textContent = `(${el.childElementCount})`;
}

// ---------- hiển thị ----------

function valueCell(raw, planned) {
  if (planned) return esc(raw);
  const t = String(raw ?? '').trim();
  return `<span class="skipval">${t === '' ? '(trống)' : esc(t)}</span>`;
}

function proxyCell(row) {
  const p = row.plan.proxy;
  if (!p) return valueCell(row.proxy, false);
  if (p.error) return `<span class="mono">${esc(row.proxy)}</span><div class="sub bad">⚠ ${esc(p.error)}</div>`;
  return `<span class="mono">${esc(row.proxy)}</span><div class="sub mono">→ ${esc(p.value)}</div>`;
}

function statusCell(id, row) {
  const res = state.results.get(id);
  const planned = KEYS.filter((k) => row.plan[k]);

  if (res) {
    const attempted = KEYS.filter((k) => res.results[k] && res.results[k].state !== 'skip');
    if (!attempted.length) return '<span class="muted">Bỏ qua (không có dữ liệu)</span>';
    const chips = attempted
      .map((k) => {
        const r = res.results[k];
        const icon = r.state === 'ok' ? '✓' : '✗';
        return `<span class="chip ${r.state}" title="${esc(r.message || '')}">${LABELS[k]} ${icon}</span>`;
      })
      .join('');
    const errors = attempted
      .filter((k) => res.results[k].state === 'error')
      .map((k) => `<div class="err-msg">${LABELS[k]}: ${esc(res.results[k].message)}</div>`)
      .join('');
    return `<div class="chips">${chips}</div>${errors}`;
  }

  if (!planned.length) return '<span class="muted">Không có gì để cập nhật</span>';
  const isRunning = state.runningId === id;
  return `<div class="chips">${planned
    .map((k) => {
      if (row.plan[k].error) return `<span class="chip error" title="${esc(row.plan[k].error)}">${LABELS[k]} ✗</span>`;
      return `<span class="chip ${isRunning ? 'running' : 'pending'}">${isRunning ? '⏳ ' : ''}${LABELS[k]}</span>`;
    })
    .join('')}</div>`;
}

function rowHtml(sheetName, row) {
  const id = rowId(sheetName, row);
  const checked = !state.unchecked.has(id);
  const cls = [checked ? '' : 'unchecked', state.runningId === id ? 'running' : ''].join(' ').trim();
  return `<tr data-id="${esc(id)}" class="${cls}">
    <td class="col-check"><input type="checkbox" class="row-check" ${checked ? 'checked' : ''} ${state.running ? 'disabled' : ''} /></td>
    <td class="muted">${row.excelRow}</td>
    <td class="mono">${esc(row.uid)}</td>
    <td>${valueCell(row.name, !!row.plan.name)}</td>
    <td>${valueCell(row.note, !!row.plan.note)}</td>
    <td>${proxyCell(row)}</td>
    <td class="status-cell">${statusCell(id, row)}</td>
  </tr>`;
}

function renderTable() {
  const sheet = currentSheet();
  const body = $('tableBody');
  if (!sheet) {
    body.innerHTML = '<tr class="empty"><td colspan="7">Chọn file Excel để hiển thị dữ liệu</td></tr>';
  } else if (!sheet.rows.length) {
    body.innerHTML = `<tr class="empty"><td colspan="7">Sheet "${esc(sheet.name)}" không có dòng dữ liệu nào (cột A trống)</td></tr>`;
  } else {
    body.innerHTML = sheet.rows.map((r) => rowHtml(sheet.name, r)).join('');
  }
  updateCheckAll();
  updateStats();
  updateModeHint();
  updateButtons();
}

function renderRow(id) {
  const sheet = currentSheet();
  if (!sheet) return;
  const row = sheet.rows.find((r) => rowId(sheet.name, r) === id);
  const tr = document.querySelector(`tr[data-id="${CSS.escape(id)}"]`);
  if (!row || !tr) return;
  tr.outerHTML = rowHtml(sheet.name, row);
}

function rowOutcome(id) {
  const res = state.results.get(id);
  if (!res) return null;
  const states = KEYS.map((k) => res.results[k] && res.results[k].state).filter((s) => s && s !== 'skip');
  if (!states.length) return 'skip';
  return states.includes('error') ? 'error' : 'ok';
}

function updateStats() {
  const sheet = currentSheet();
  const rows = currentRows();
  let selected = 0;
  let ok = 0;
  let err = 0;
  let skip = 0;
  for (const r of rows) {
    const id = rowId(sheet.name, r);
    if (!state.unchecked.has(id)) selected++;
    const o = rowOutcome(id);
    if (o === 'ok') ok++;
    else if (o === 'error') err++;
    else if (o === 'skip') skip++;
  }
  $('statTotal').textContent = rows.length;
  $('statSelected').textContent = selected;
  $('statOk').textContent = ok;
  $('statErr').textContent = err;
  $('statSkip').textContent = skip;
  $('btnRun').textContent = state.running ? '⏳ Đang cập nhật…' : `▶ Bắt đầu cập nhật (${selected})`;
  const total = state.queue.length;
  $('progressBar').style.width = total ? `${Math.round((state.done / total) * 100)}%` : '0%';
}

function updateCheckAll() {
  const sheet = currentSheet();
  const rows = currentRows();
  const box = $('checkAll');
  const checkedCount = rows.filter((r) => !state.unchecked.has(rowId(sheet.name, r))).length;
  box.checked = rows.length > 0 && checkedCount === rows.length;
  box.indeterminate = checkedCount > 0 && checkedCount < rows.length;
  box.disabled = state.running || !rows.length;
}

function updateModeHint() {
  const rows = currentRows();
  const hint = $('modeHint');
  if (!rows.length) return hint.classList.add('hidden');
  const localCount = rows.filter((r) => /^local-/i.test(r.uid)).length;
  const mode = getMode();
  let msg = '';
  if (mode === 'cloud' && localCount > 0) {
    msg = `⚠ Có ${localCount}/${rows.length} UID bắt đầu bằng "local-" (profile Local) nhưng đang chọn chế độ Cloud.`;
  } else if (mode === 'local' && localCount < rows.length) {
    msg = `⚠ Có ${rows.length - localCount}/${rows.length} UID không bắt đầu bằng "local-" nhưng đang chọn chế độ Local.`;
  }
  hint.textContent = msg;
  hint.classList.toggle('hidden', !msg);
}

function updateButtons() {
  const hasRows = currentRows().length > 0;
  const hasFile = !!state.file;
  const sheet = currentSheet();
  const selected = hasRows && currentRows().some((r) => !state.unchecked.has(rowId(sheet.name, r)));
  const anyResult = hasRows && currentRows().some((r) => state.results.has(rowId(sheet.name, r)));
  const anyRetry =
    anyResult && currentRows().some((r) => [null, 'error'].includes(rowOutcome(rowId(sheet.name, r))));

  $('btnOpen').disabled = state.running;
  $('btnReload').disabled = state.running || !hasFile;
  $('sheetSelect').disabled = state.running || !hasFile;
  $('btnRun').disabled = state.running || !selected;
  $('btnStop').disabled = !state.running;
  $('btnSelectFailed').disabled = state.running || !anyRetry;
  $('btnExport').disabled = state.running || !anyResult;
  document.querySelectorAll('input[name="mode"]').forEach((el) => (el.disabled = state.running));
  $('baseUrl').disabled = state.running;
  $('delayMs').disabled = state.running;
}

// ---------- tải file ----------

function applyFile(data, keepSheetName) {
  state.file = data;
  state.results.clear();
  state.unchecked.clear();
  state.queue = [];
  state.done = 0;

  const select = $('sheetSelect');
  select.innerHTML = data.sheets
    .map((s, i) => `<option value="${i}">${esc(s.name)} (${s.rows.length} dòng)</option>`)
    .join('');
  let idx = keepSheetName ? data.sheets.findIndex((s) => s.name === keepSheetName) : -1;
  if (idx < 0) idx = Math.max(0, data.sheets.findIndex((s) => s.rows.length > 0));
  state.sheetIndex = idx;
  select.value = String(idx);

  $('fileName').textContent = fileNameOf(data.filePath);
  $('fileName').title = data.filePath;
  $('fileName').className = '';
  renderTable();
  saveSettings();
  log(`Đã tải ${fileNameOf(data.filePath)}: ${data.sheets.map((s) => `${s.name}=${s.rows.length} dòng`).join(', ')}`);
}

// Mở lại file đã dùng lần trước (đọc lại nội dung mới nhất từ đĩa).
async function restoreLastFile(filePath, sheetName) {
  if (!filePath) return;
  try {
    applyFile(await getApi().reloadExcel(filePath), sheetName);
  } catch (err) {
    $('fileName').textContent = `Không mở được file lần trước: ${fileNameOf(filePath)}`;
    $('fileName').title = `${filePath}\n${err.message}`;
    $('fileName').className = 'missing';
    log(`Không mở được file lần trước (${filePath}): ${err.message}`, 'error');
  }
}

async function openFile() {
  try {
    const data = await getApi().openExcel();
    if (data) applyFile(data);
  } catch (err) {
    alert(`Không đọc được file Excel:\n${err.message}`);
  }
}

async function reloadFile() {
  if (!state.file) return;
  if (state.results.size && !confirm('Tải lại sẽ xóa trạng thái đã chạy trên bảng. Tiếp tục?')) return;
  try {
    const sheet = currentSheet();
    applyFile(await getApi().reloadExcel(state.file.filePath), sheet && sheet.name);
  } catch (err) {
    alert(`Không đọc được file Excel:\n${err.message}`);
  }
}

// ---------- chạy API ----------

function bindProgress() {
  if (state.progressBound) return;
  state.progressBound = true;
  getApi().onProgress(onProgress);
}

// Kết quả 1 dòng: đến qua sự kiện progress hoặc từ danh sách trả về khi chạy xong
// (IPC không đảm bảo thứ tự giữa hai kênh nên phải bỏ qua bản trùng).
function onProgress({ runToken, id, results, summary }) {
  if (runToken !== state.runToken || state.results.has(id)) return;
  state.results.set(id, { results, summary });
  state.done++;
  const next = state.queue[state.done];
  const prevRunning = state.runningId;
  state.runningId = next ? next.id : null;
  renderRow(id);
  if (prevRunning && prevRunning !== id) renderRow(prevRunning);
  if (state.runningId) renderRow(state.runningId);

  const row = state.queue.find((q) => q.id === id);
  const label = row ? `Dòng ${row.excelRow} · ${row.uid}` : id;
  const attempted = KEYS.filter((k) => results[k] && results[k].state !== 'skip');
  if (!attempted.length) log(`${label} → bỏ qua (không có dữ liệu)`);
  for (const k of attempted) {
    const r = results[k];
    log(`${label} · ${LABELS[k]} → ${r.state === 'ok' ? 'OK' : 'LỖI'}: ${r.message}`, r.state);
  }
  updateStats();
  updateCheckAll();
  updateButtons();
}

async function run() {
  const sheet = currentSheet();
  if (!sheet || state.running) return;
  const rows = sheet.rows
    .map((r) => ({ ...r, id: rowId(sheet.name, r) }))
    .filter((r) => !state.unchecked.has(r.id));
  if (!rows.length) return;

  const mode = getMode();
  const baseUrl = $('baseUrl').value.trim();
  const delayMs = Math.max(0, Number($('delayMs').value) || 0);
  saveSettings();

  let apiRef;
  try {
    apiRef = getApi();
    bindProgress();
  } catch (err) {
    return alert(err.message);
  }

  for (const r of rows) state.results.delete(r.id);
  state.runToken = `${Date.now()}-${Math.random()}`;
  state.queue = rows;
  state.done = 0;
  state.running = true;
  state.runningId = rows[0].id;
  renderTable();
  log(`▶ Bắt đầu: ${rows.length} dòng · sheet ${sheet.name} · ${mode.toUpperCase()} · ${baseUrl}`);

  try {
    const res = await apiRef.run({
      runToken: state.runToken,
      rows: rows.map(({ id, excelRow, uid, name, note, proxy, plan }) => ({ id, excelRow, uid, name, note, proxy, plan })),
      baseUrl,
      isLocal: mode === 'local',
      delayMs,
    });
    for (const item of (res && res.items) || []) onProgress({ ...item, runToken: state.runToken });
    log(res && res.stopped ? `■ Đã dừng sau ${res.done} dòng` : `✔ Hoàn tất ${res ? res.done : state.done} dòng`);
  } catch (err) {
    log(`Lỗi khi chạy: ${err.message}`, 'error');
    alert(`Lỗi khi chạy:\n${err.message}`);
  } finally {
    state.running = false;
    state.runningId = null;
    renderTable();
  }
}

async function stop() {
  try {
    await getApi().stop();
    log('Đang dừng… (sẽ dừng sau dòng hiện tại)');
    $('btnStop').disabled = true;
  } catch (err) {
    alert(err.message);
  }
}

// Tick các dòng lỗi hoặc chưa chạy, bỏ tick các dòng đã thành công / bỏ qua.
function selectFailed() {
  const sheet = currentSheet();
  if (!sheet) return;
  for (const r of sheet.rows) {
    const id = rowId(sheet.name, r);
    if ([null, 'error'].includes(rowOutcome(id))) state.unchecked.delete(id);
    else state.unchecked.add(id);
  }
  renderTable();
}

async function exportResults() {
  const sheet = currentSheet();
  if (!sheet) return;
  const statuses = {};
  for (const r of sheet.rows) {
    const res = state.results.get(rowId(sheet.name, r));
    if (res) statuses[r.excelRow] = res.summary;
  }
  try {
    const saved = await getApi().exportResults({ sourcePath: state.file.filePath, sheetName: sheet.name, statuses });
    if (saved) {
      log(`💾 Đã lưu kết quả: ${saved}`);
      alert(`Đã lưu kết quả vào:\n${saved}`);
    }
  } catch (err) {
    alert(`Không lưu được file:\n${err.message}`);
  }
}

// ---------- sự kiện ----------

function init() {
  const settings = loadSettings();

  $('btnOpen').addEventListener('click', openFile);
  $('btnReload').addEventListener('click', reloadFile);
  $('btnRun').addEventListener('click', run);
  $('btnStop').addEventListener('click', stop);
  $('btnSelectFailed').addEventListener('click', selectFailed);
  $('btnExport').addEventListener('click', exportResults);

  $('sheetSelect').addEventListener('change', (e) => {
    state.sheetIndex = Number(e.target.value);
    renderTable();
    saveSettings();
  });

  document.querySelectorAll('input[name="mode"]').forEach((el) =>
    el.addEventListener('change', () => {
      updateModeHint();
      saveSettings();
    }),
  );
  $('baseUrl').addEventListener('change', saveSettings);
  $('delayMs').addEventListener('change', saveSettings);

  $('checkAll').addEventListener('change', (e) => {
    const sheet = currentSheet();
    if (!sheet) return;
    for (const r of sheet.rows) {
      const id = rowId(sheet.name, r);
      if (e.target.checked) state.unchecked.delete(id);
      else state.unchecked.add(id);
    }
    renderTable();
  });

  $('tableBody').addEventListener('change', (e) => {
    if (!e.target.classList.contains('row-check')) return;
    const tr = e.target.closest('tr');
    const id = tr.dataset.id;
    if (e.target.checked) state.unchecked.delete(id);
    else state.unchecked.add(id);
    tr.classList.toggle('unchecked', !e.target.checked);
    updateCheckAll();
    updateStats();
    updateButtons();
  });

  renderTable();
  if (window.api) restoreLastFile(settings.lastFile, settings.lastSheet);
}

document.addEventListener('DOMContentLoaded', init);
