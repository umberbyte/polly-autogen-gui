const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('spreadsheetAPI', {
  onInit: (callback) => {
    ipcRenderer.on('spreadsheet:init', (event, data) => callback(data));
  },
  save: (payload) => ipcRenderer.invoke('spreadsheet:save', payload),
  generateOne: (payload) => ipcRenderer.invoke('spreadsheet:generate-one', payload),
  batchGenerate: (payload) => ipcRenderer.invoke('spreadsheet:batch-generate', payload),
  onBatchProgress: (callback) => {
    ipcRenderer.on('spreadsheet:batch-progress', (event, data) => callback(data));
  },
});
