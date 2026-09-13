const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('pollyAutogen', {
  version: process.versions.electron,
  selectFile: () => ipcRenderer.invoke('dialog:select-file'),
  selectDirectory: () => ipcRenderer.invoke('dialog:select-directory'),
  selectExcelFile: () => ipcRenderer.invoke('dialog:select-excel-file'),
  selectWorkingFolder: () => ipcRenderer.invoke('dialog:select-working-folder'),
  startExcelImport: (excelPath) => ipcRenderer.invoke('excel:start-import', { excelPath }),
  startWorkFolderReview: (folderPath) => ipcRenderer.invoke('workfolder:start-review', { folderPath }),
  listVoices: (region) => ipcRenderer.invoke('polly:list-voices', { region }),
  loadVoiceSettings: () => ipcRenderer.invoke('settings:load-voice-settings'),
  saveVoiceSettings: (settings) => ipcRenderer.invoke('settings:save-voice-settings', settings),
  convert: (payload) => ipcRenderer.invoke('polly:convert', payload),
  selectPptxFile: () => ipcRenderer.invoke('dialog:select-pptx-file'),
  preparePptxWorkingCopy: (pptxPath) => ipcRenderer.invoke('pptx:prepare-working-copy', { pptxPath }),
  embedAudioToPptx: (payload) => ipcRenderer.invoke('pptx:embed-audio', payload),
  onPptxEmbedProgress: (callback) => {
    ipcRenderer.on('pptx:embed-progress', (event, data) => callback(data));
  },
  onProgress: (callback) => {
    const listener = (event, data) => callback(data);
    ipcRenderer.on('polly:progress', listener);
    return () => ipcRenderer.removeListener('polly:progress', listener);
  },
});
