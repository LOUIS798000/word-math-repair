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
  selectDirectory: () => ipcRenderer.invoke('dialog:selectDirectory'),
  analyzeDocx: (filePath, options) => ipcRenderer.invoke('engine:analyze', filePath, options),
  repairDocx: (filePath, outputPath, options, exportPdf = false) => ipcRenderer.invoke('engine:repair', filePath, outputPath, options, exportPdf),
  repairBatch: (filePaths, outputDir, options, exportPdf = false) => ipcRenderer.invoke('engine:repairBatch', filePaths, outputDir, options, exportPdf),
  onBatchProgress: (callback) => {
    const listener = (_event, progress) => callback(progress);
    ipcRenderer.on('batch:progress', listener);
    return () => ipcRenderer.removeListener('batch:progress', listener);
  },
  renderPages: (filePath) => ipcRenderer.invoke('engine:renderPages', filePath),
  openInWord: (filePath) => ipcRenderer.invoke('app:openInWord', filePath),
  openPath: (path) => ipcRenderer.invoke('shell:openPath', path),
  getDefaultSamplePath: () => ipcRenderer.invoke('app:getDefaultSamplePath'),
  getLaunchFile: () => ipcRenderer.invoke('app:getLaunchFile'),
  onOpenFile: (callback) => {
    const listener = (_event, filePath) => callback(filePath);
    ipcRenderer.on('app:openFile', listener);
    return () => ipcRenderer.removeListener('app:openFile', listener);
  }
});
