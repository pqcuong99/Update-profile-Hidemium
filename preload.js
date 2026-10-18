const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('api', {
  openExcel: () => ipcRenderer.invoke('excel:open'),
  reloadExcel: (filePath) => ipcRenderer.invoke('excel:reload', filePath),
  exportResults: (payload) => ipcRenderer.invoke('excel:export', payload),
  run: (payload) => ipcRenderer.invoke('api:run', payload),
  stop: () => ipcRenderer.invoke('api:stop'),
  onProgress: (callback) => {
    const listener = (_e, data) => callback(data);
    ipcRenderer.on('api:progress', listener);
    return () => ipcRenderer.removeListener('api:progress', listener);
  },
});
