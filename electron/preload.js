const { contextBridge, ipcRenderer } = require('electron')

// 安全地暴露API给渲染进程
contextBridge.exposeInMainWorld('electronAPI', {
  selectFile: () => ipcRenderer.invoke('select-file'),
  saveFile: (defaultName, content) => ipcRenderer.invoke('save-file', { defaultName, content }),
  saveBuffer: (filePath, buffer) => ipcRenderer.invoke('save-buffer', { filePath, buffer }),
  readFile: (filePath) => ipcRenderer.invoke('read-file', filePath),
  getAppPath: () => ipcRenderer.invoke('get-app-path'),
  // PDF生成接口
  generatePdf: (htmlContent, options) => ipcRenderer.invoke('generate-pdf', { htmlContent, options }),
  // 打开PDF预览窗口
  openPdfPreview: (htmlContent) => ipcRenderer.invoke('open-pdf-preview', { htmlContent })
})