const { contextBridge, ipcRenderer, webUtils } = require('electron');

contextBridge.exposeInMainWorld('api', {
  getPathForFile: (file) => {
    try {
      if (webUtils && typeof webUtils.getPathForFile === 'function') {
        return webUtils.getPathForFile(file);
      }
    } catch (_) {}
    return file.path || '';
  },
  openFileDialog: (allowMultiple = false) => ipcRenderer.invoke('dialog:openFile', allowMultiple),
  saveFileDialog: (defaultName) => ipcRenderer.invoke('dialog:saveFile', defaultName),
  analyzeDocx: (filePath, options) => ipcRenderer.invoke('engine:analyze', filePath, options),
  repairDocx: (filePath, outputPath, options, exportPdf = false) => ipcRenderer.invoke('engine:repair', filePath, outputPath, options, exportPdf),
  renderPages: (filePath) => ipcRenderer.invoke('engine:renderPages', filePath),
  openInWord: (filePath) => ipcRenderer.invoke('app:openInWord', filePath),
  openPath: (path) => ipcRenderer.invoke('shell:openPath', path),
  getDefaultSamplePath: () => ipcRenderer.invoke('app:getDefaultSamplePath')
});
