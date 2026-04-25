# Electron 开发指南

## 快速开始

### 1. 安装依赖

```bash
# 根目录
npm install

# 前端
cd frontend && npm install

# 后端
cd ../backend && pip install -r requirements.txt
```

### 2. 启动开发环境

```bash
# 方式一：一键启动（推荐）
npm run dev

# 方式二：分别启动
# 终端 1: 启动前端
npm run dev:frontend

# 终端 2: 启动后端
npm run dev:backend

# 终端 3: 启动 Electron
npm run dev:electron
```

### 3. 访问应用

- Electron 窗口会自动打开
- 前端开发服务器: http://localhost:5173
- 后端 API: http://localhost:8000
- API 文档: http://localhost:8000/docs

## 项目结构

```
customs-optimizer/
├── electron/                    # Electron 主进程
│   ├── main.js                  # 主进程入口
│   ├── preload.js               # 预加载脚本（安全桥接）
│   └── python-bridge.js         # Python 后端管理
├── frontend/                    # React 前端
│   ├── src/
│   │   ├── components/          # 组件
│   │   ├── pages/               # 页面
│   │   ├── services/            # 服务层
│   │   └── App.tsx
│   └── package.json
├── backend/                     # Python 后端
│   ├── app/
│   │   ├── main.py              # FastAPI 入口
│   │   ├── api/                 # API 路由
│   │   ├── services/            # 业务逻辑
│   │   ├── models/              # 数据模型
│   │   └── utils/               # 工具函数
│   └── requirements.txt
├── build/                       # 构建配置
│   └── electron-builder.yml
├── resources/                   # 资源文件
│   └── icon.ico
└── package.json                 # 根配置
```

## 核心概念

### Electron 架构

```
┌─────────────────────────────────────────┐
│         Electron 主进程 (main.js)        │
│  - 窗口管理                              │
│  - Python 后端启动/停止                  │
│  - 文件系统访问                          │
│  - 本地存储                              │
└──────────────┬──────────────────────────┘
               │ IPC 通信
┌──────────────▼──────────────────────────┐
│      Electron 渲染进程 (React)           │
│  - 用户界面                              │
│  - 通过 preload.js 调用主进程 API       │
└──────────────┬──────────────────────────┘
               │ HTTP
┌──────────────▼──────────────────────────┐
│         Python 后端 (FastAPI)            │
│  - 业务逻辑                              │
│  - 数据处理                              │
└─────────────────────────────────────────┘
```

### IPC 通信

主进程和渲染进程通过 IPC（进程间通信）交互：

**主进程 (main.js)**
```javascript
ipcMain.handle('get-settings', () => {
  return store.get('settings', {});
});
```

**预加载脚本 (preload.js)**
```javascript
contextBridge.exposeInMainWorld('electronAPI', {
  getSettings: () => ipcRenderer.invoke('get-settings')
});
```

**渲染进程 (React)**
```javascript
const settings = await window.electronAPI.getSettings();
```

## 开发技巧

### 1. 调试

**主进程调试**
```bash
# 启动时添加调试参数
electron --inspect=5858 .
```

**渲染进程调试**
- 开发模式下会自动打开 DevTools
- 或按 `Ctrl+Shift+I` (Windows/Linux) / `Cmd+Option+I` (Mac)

**Python 后端调试**
```bash
# 使用 --reload 自动重载
python -m uvicorn app.main:app --reload
```

### 2. 热重载

- **前端**: Vite 自动热重载
- **后端**: uvicorn --reload 自动重载
- **Electron**: 需要手动重启（或使用 electron-reload）

### 3. 环境变量

开发环境通过 `NODE_ENV=development` 区分：

```javascript
if (process.env.NODE_ENV === 'development') {
  // 开发环境：加载 localhost
  mainWindow.loadURL('http://localhost:5173');
} else {
  // 生产环境：加载打包后的文件
  mainWindow.loadFile('frontend/dist/index.html');
}
```

### 4. 本地存储

使用 electron-store 存储配置：

```javascript
const Store = require('electron-store');
const store = new Store();

// 保存
store.set('settings', { theme: 'dark' });

// 读取
const settings = store.get('settings');

// 删除
store.delete('settings');
```

存储位置：
- Windows: `%APPDATA%/customs-optimizer/config.json`
- macOS: `~/Library/Application Support/customs-optimizer/config.json`
- Linux: `~/.config/customs-optimizer/config.json`

## 常见问题

### Q1: Python 后端启动失败

**原因**：
- Python 未安装或版本不对
- 依赖未安装
- 端口被占用

**解决**：
```bash
# 检查 Python 版本
python --version  # 需要 >= 3.10

# 安装依赖
cd backend
pip install -r requirements.txt

# 检查端口占用
# Windows
netstat -ano | findstr 8000
# Mac/Linux
lsof -i :8000
```

### Q2: 前端无法连接后端

**原因**：
- 后端未启动
- CORS 配置问题
- 端口不匹配

**解决**：
1. 确认后端已启动：访问 http://localhost:8000/health
2. 检查 CORS 配置（backend/app/main.py）
3. 检查前端 API 地址配置

### Q3: Electron 窗口空白

**原因**：
- 前端未构建或未启动
- 路径配置错误

**解决**：
1. 开发模式：确认前端开发服务器已启动
2. 生产模式：确认前端已构建 `npm run build:frontend`
3. 打开 DevTools 查看错误信息

### Q4: 打包后无法启动

**原因**：
- Python 运行时未打包
- 路径配置错误
- 依赖缺失

**解决**：
1. 检查 `build/electron-builder.yml` 配置
2. 确认 `extraResources` 包含 backend 目录
3. 测试打包：`npm run pack`（不生成安装包）

## 构建打包

### 开发构建（测试）

```bash
# 只打包不生成安装包
npm run pack

# 输出到 dist/win-unpacked/ 或 dist/mac/
```

### 生产构建

```bash
# 1. 构建前端
npm run build:frontend

# 2. 打包 Electron（包含安装包）
npm run build:win      # Windows (NSIS + MSI)
npm run build:mac      # macOS (DMG)
npm run build:linux    # Linux (AppImage)

# 输出到 dist/
```

### 打包配置

编辑 `build/electron-builder.yml`：

```yaml
# 应用信息
appId: com.customs.optimizer
productName: 清关优化系统

# 包含的文件
files:
  - electron/**/*
  - frontend/dist/**/*
  - package.json

# 额外资源（会复制到 resources 目录）
extraResources:
  - from: backend
    to: backend
```

## 发布流程

### 1. 版本更新

```bash
# 更新版本号
npm version patch  # 1.0.0 -> 1.0.1
npm version minor  # 1.0.0 -> 1.1.0
npm version major  # 1.0.0 -> 2.0.0
```

### 2. 构建安装包

```bash
npm run dist:win
```

### 3. 测试安装包

1. 在干净的环境中安装
2. 测试所有功能
3. 检查更新机制

### 4. 发布

1. 上传安装包到服务器
2. 更新下载链接
3. 通知用户

## 性能优化

### 1. 减小安装包体积

- 使用 `asar` 打包（默认启用）
- 排除不必要的文件
- 压缩资源文件

### 2. 加快启动速度

- 延迟加载非关键模块
- 优化 Python 后端启动
- 使用 splash screen

### 3. 降低内存占用

- 及时释放不用的资源
- 优化图片和资源加载
- 使用 Web Workers 处理耗时任务

## 安全建议

### 1. 渲染进程安全

```javascript
// ✅ 正确：使用 contextIsolation
webPreferences: {
  nodeIntegration: false,
  contextIsolation: true,
  preload: path.join(__dirname, 'preload.js')
}

// ❌ 错误：不要启用 nodeIntegration
webPreferences: {
  nodeIntegration: true  // 危险！
}
```

### 2. IPC 安全

```javascript
// ✅ 正确：验证输入
ipcMain.handle('save-file', async (event, data) => {
  if (!isValidData(data)) {
    throw new Error('Invalid data');
  }
  // ...
});

// ❌ 错误：直接使用用户输入
ipcMain.handle('exec', (event, cmd) => {
  exec(cmd);  // 危险！
});
```

### 3. 内容安全策略

在 HTML 中添加 CSP：

```html
<meta http-equiv="Content-Security-Policy"
      content="default-src 'self'; script-src 'self'">
```

## 下一步

1. 添加更多 API 路由
2. 实现文件上传功能
3. 添加数据处理逻辑
4. 完善错误处理
5. 添加日志系统
6. 实现自动更新

## 参考资源

- [Electron 官方文档](https://www.electronjs.org/docs)
- [FastAPI 文档](https://fastapi.tiangolo.com/)
- [React 文档](https://react.dev/)
- [electron-builder 文档](https://www.electron.build/)
