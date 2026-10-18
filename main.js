const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const core = require('./src/core');

let mainWindow;
let stopRequested = false;
let running = false;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1280,
    height: 860,
    minWidth: 900,
    minHeight: 600,
    title: 'Excel API Updater',
    icon: path.join(__dirname, 'images', 'app.ico'),
    backgroundColor: '#f4f5fb',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  });
  mainWindow.loadFile(path.join(__dirname, 'renderer', 'index.html'));
}

ipcMain.handle('excel:open', async () => {
  const res = await dialog.showOpenDialog(mainWindow, {
    title: 'Chọn file Excel',
    properties: ['openFile'],
    filters: [
      { name: 'Excel', extensions: ['xlsx', 'xlsm', 'xls', 'csv'] },
      { name: 'Tất cả', extensions: ['*'] },
    ],
  });
  if (res.canceled || !res.filePaths.length) return null;
  return core.readExcel(res.filePaths[0]);
});

ipcMain.handle('excel:reload', async (_e, filePath) => core.readExcel(filePath));

ipcMain.handle('excel:export', async (_e, { sourcePath, sheetName, statuses }) => {
  const parsed = path.parse(sourcePath);
  const res = await dialog.showSaveDialog(mainWindow, {
    title: 'Lưu kết quả',
    defaultPath: path.join(parsed.dir, `${parsed.name}_ket_qua.xlsx`),
    filters: [{ name: 'Excel', extensions: ['xlsx'] }],
  });
  if (res.canceled || !res.filePath) return null;
  return core.exportResults({ sourcePath, sheetName, statuses, savePath: res.filePath });
});

ipcMain.handle('api:run', async (event, { runToken, rows, baseUrl, isLocal, delayMs }) => {
  if (running) throw new Error('Đang chạy, vui lòng đợi hoặc bấm Dừng.');
  running = true;
  stopRequested = false;
  const items = [];
  try {
    const res = await core.runUpdates(rows, {
      baseUrl,
      isLocal,
      delayMs,
      shouldStop: () => stopRequested,
      onProgress: (p) => {
        const item = { id: p.row.id, results: p.results, summary: p.summary };
        items.push(item);
        if (!event.sender.isDestroyed()) event.sender.send('api:progress', { ...item, runToken });
      },
    });
    // Trả kèm toàn bộ kết quả: sự kiện progress có thể đến renderer sau phản hồi này.
    return { ...res, items };
  } finally {
    running = false;
  }
});

ipcMain.handle('api:stop', () => {
  stopRequested = true;
  return true;
});

app.whenReady().then(() => {
  createWindow();
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});
