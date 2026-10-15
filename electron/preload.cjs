const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  savePdf: (options) => ipcRenderer.invoke('save-pdf', options),
  exportFile: (options) => ipcRenderer.invoke('export-file', options),
  importFile: (options) => ipcRenderer.invoke('import-file', options),
  isElectron: true,
});

