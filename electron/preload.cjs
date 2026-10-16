const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  savePdf: (options) => ipcRenderer.invoke('save-pdf', options),
  exportFile: (options) => ipcRenderer.invoke('export-file', options),
  importFile: (options) => ipcRenderer.invoke('import-file', options),
  saveBackup: (options) => ipcRenderer.invoke('save-backup', options),
  listBackups: () => ipcRenderer.invoke('list-backups'),
  loadBackup: (options) => ipcRenderer.invoke('load-backup', options),
  fetchMarketPrices: () => ipcRenderer.invoke('fetch-market-prices'),
  isElectron: true,
});

