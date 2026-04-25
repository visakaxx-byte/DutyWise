const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const Store = require('electron-store');
const PythonBridge = require('./python-bridge');

const store = new Store();
let mainWindow;
let pythonBridge;

// 创建主窗口
function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 800,
    minWidth: 1000,
    minHeight: 600,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      nodeIntegration: false,
      contextIsolation: true,
      sandbox: false
    },
    icon: path.join(__dirname, '../resources/icon.ico')
  });

  // 加载前端
  if (process.env.NODE_ENV === 'development') {
    mainWindow.loadURL('http://localhost:5173');
    mainWindow.webContents.openDevTools();
  } else {
    mainWindow.loadFile(path.join(__dirname, '../frontend/dist/index.html'));
  }

  // 窗口关闭事件
  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

// 应用启动
app.whenReady().then(async () => {
  // 启动 Python 后端
  pythonBridge = new PythonBridge();

  try {
    await pythonBridge.start();
    console.log('Python 后端启动成功');
  } catch (error) {
    console.error('Python 后端启动失败:', error);
    dialog.showErrorBox('启动失败', `无法启动后端服务: ${error.message}`);
    app.quit();
    return;
  }

  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

// 应用退出
app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('quit', () => {
  if (pythonBridge) {
    pythonBridge.stop();
  }
});

// IPC 通信处理

// 获取设置
ipcMain.handle('get-settings', () => {
  return store.get('settings', {
    apiUrl: 'http://127.0.0.1:8000',
    theme: 'light',
    language: 'zh-CN'
  });
});

// 保存设置
ipcMain.handle('save-settings', (event, settings) => {
  try {
    store.set('settings', settings);
    return { success: true };
  } catch (error) {
    return { success: false, error: error.message };
  }
});

// 选择文件
ipcMain.handle('select-files', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ['openFile', 'multiSelections'],
    filters: [
      { name: 'Excel 文件', extensions: ['xlsx', 'xls'] },
      { name: '所有文件', extensions: ['*'] }
    ]
  });

  if (result.canceled) {
    return [];
  }

  return result.filePaths;
});

// 选择文件夹
ipcMain.handle('select-folder', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ['openDirectory']
  });

  if (result.canceled) {
    return null;
  }

  return result.filePaths[0];
});

// 保存文件对话框
ipcMain.handle('save-file', async (event, options) => {
  const result = await dialog.showSaveDialog(mainWindow, {
    defaultPath: options.defaultPath || 'output.xlsx',
    filters: options.filters || [
      { name: 'Excel 文件', extensions: ['xlsx'] },
      { name: '所有文件', extensions: ['*'] }
    ]
  });

  if (result.canceled) {
    return null;
  }

  return result.filePath;
});

// 打开外部链接
ipcMain.handle('open-external', async (event, url) => {
  const { shell } = require('electron');
  await shell.openExternal(url);
});

// 获取应用信息
ipcMain.handle('get-app-info', () => {
  return {
    name: app.getName(),
    version: app.getVersion(),
    platform: process.platform,
    arch: process.arch
  };
});

// 获取后端状态
ipcMain.handle('get-backend-status', () => {
  return pythonBridge ? pythonBridge.getStatus() : { running: false };
});

// 重启后端
ipcMain.handle('restart-backend', async () => {
  if (pythonBridge) {
    pythonBridge.stop();
    await pythonBridge.start();
    return { success: true };
  }
  return { success: false, error: 'Python bridge not initialized' };
});
