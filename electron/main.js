const { app, BrowserWindow, ipcMain, dialog } = require('electron')
const path = require('path')
const fs = require('fs')

// 禁用硬件加速以提高某些系统的兼容性
// app.disableHardwareAcceleration()

// 存储主窗口引用
let mainWindow = null

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 800,
    minWidth: 900,
    minHeight: 600,
    icon: path.resolve(__dirname, '../build/icon.ico'),
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: path.join(__dirname, 'preload.js')
    },
    titleBarStyle: 'hiddenInset',
    frame: process.platform === 'darwin' ? true : true
  })

  // 开发模式加载Vite服务器
  if (process.env.NODE_ENV === 'development' || !app.isPackaged) {
    mainWindow.loadURL('http://localhost:5173')
    mainWindow.webContents.openDevTools()
  } else {
    // 生产模式加载打包后的文件
    mainWindow.loadFile(path.join(__dirname, '../dist/index.html'))
  }
}

// IPC处理：选择文件
ipcMain.handle('select-file', async () => {
  const result = await dialog.showOpenDialog({
    filters: [
      { name: 'Markdown', extensions: ['md', 'markdown', 'txt'] }
    ],
    properties: ['openFile']
  })

  if (result.canceled || result.filePaths.length === 0) {
    return null
  }

  const filePath = result.filePaths[0]
  const content = fs.readFileSync(filePath, 'utf-8')
  return {
    path: filePath,
    name: path.basename(filePath),
    content: content
  }
})

// IPC处理：保存文件
ipcMain.handle('save-file', async (event, { defaultName, content }) => {
  const result = await dialog.showSaveDialog({
    filters: [
      { name: 'Word Document', extensions: ['docx'] },
      { name: 'PDF', extensions: ['pdf'] }
    ],
    defaultPath: defaultName
  })

  if (result.canceled || !result.filePath) {
    return null
  }

  return result.filePath
})

// IPC处理：保存Buffer数据
ipcMain.handle('save-buffer', async (event, { filePath, buffer }) => {
  fs.writeFileSync(filePath, Buffer.from(buffer))
  return true
})

// IPC处理：读取文件
ipcMain.handle('read-file', async (event, filePath) => {
  const content = fs.readFileSync(filePath, 'utf-8')
  return content
})

// IPC处理：获取应用路径
ipcMain.handle('get-app-path', () => {
  return app.getPath('userData')
})

// IPC处理：生成PDF
ipcMain.handle('generate-pdf', async (event, { htmlContent, options }) => {
  try {
    // 创建一个隐藏窗口来渲染HTML并生成PDF
    const pdfWindow = new BrowserWindow({
      width: 800,
      height: 600,
      show: false, // 隐藏窗口
      webPreferences: {
        nodeIntegration: false,
        contextIsolation: true
      }
    })

    // 加载HTML内容
    pdfWindow.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(htmlContent)}`)

    // 等待页面加载完成
    await new Promise((resolve, reject) => {
      pdfWindow.webContents.on('did-finish-load', resolve)
      pdfWindow.webContents.on('did-fail-load', (event, errorCode, errorDescription) => {
        reject(new Error(`页面加载失败: ${errorDescription}`))
      })
      // 设置超时
      setTimeout(() => reject(new Error('页面加载超时')), 30000)
    })

    // 等待MathJax渲染完成
    // MathJax需要从CDN加载并渲染所有公式，这可能需要较长时间
    await new Promise(resolve => setTimeout(resolve, 5000))

    // 尝试检查MathJax是否渲染完成
    try {
      const isMathJaxReady = await pdfWindow.webContents.executeJavaScript(`
        (function() {
          if (window.mathJaxReady === true) return true;
          if (window.MathJax && window.MathJax.typesetPromise) return 'rendering';
          return false;
        })()
      `)
      if (isMathJaxReady === 'rendering') {
        // 等待额外时间让渲染完成
        await new Promise(resolve => setTimeout(resolve, 3000))
      }
    } catch (e) {
      // 忽略错误，继续生成PDF
      console.log('MathJax检查失败:', e.message)
    }

    // 生成PDF
    const pdfData = await pdfWindow.webContents.printToPDF({
      printBackground: true,
      pageSize: options.pageSize || 'A4',
      margins: {
        top: options.marginTop || 2.5, // cm
        bottom: options.marginBottom || 2.5,
        left: options.marginLeft || 2,
        right: options.marginRight || 2
      }
    })

    // 关闭隐藏窗口
    pdfWindow.close()

    // 返回PDF数据（转换为数组以便IPC传输）
    return Array.from(new Uint8Array(pdfData))
  } catch (error) {
    console.error('PDF生成错误:', error)
    throw error
  }
})

// IPC处理：打开PDF预览窗口
ipcMain.handle('open-pdf-preview', async (event, { htmlContent }) => {
  try {
    // 创建预览窗口
    const previewWindow = new BrowserWindow({
      width: 800,
      height: 600,
      title: 'PDF预览',
      webPreferences: {
        nodeIntegration: false,
        contextIsolation: true
      }
    })

    // 加载HTML内容
    previewWindow.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(htmlContent)}`)

    // 页面加载完成后打开打印对话框
    previewWindow.webContents.on('did-finish-load', () => {
      // 等待MathJax渲染
      setTimeout(() => {
        previewWindow.webContents.print({}, (success, errorType) => {
          if (!success) {
            console.log('打印失败:', errorType)
          }
        })
      }, 2000)
    })

    return true
  } catch (error) {
    console.error('PDF预览错误:', error)
    throw error
  }
})

app.whenReady().then(() => {
  createWindow()

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow()
    }
  })
})

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit()
  }
})