# DutyWise — 清关优化系统

桌面应用版本 - 基于 Electron + React + FastAPI

## 项目结构

```
DutyWise/
├── electron/              # Electron 主进程
├── electron/              # Electron 主进程
│   ├── main.js           # 主进程入口
│   ├── preload.js        # 预加载脚本
│   └── python-bridge.js  # Python 桥接
├── frontend/             # React 前端
├── backend/              # Python 后端
├── build/                # 构建配置
│   └── electron-builder.yml
├── resources/            # 资源文件
│   └── icon.ico
└── package.json          # 根配置
```

## 开发环境

### 前置要求

- Node.js >= 18.0.0
- Python >= 3.10
- npm >= 9.0.0

### 安装依赖

```bash
# 安装根目录依赖
npm install

# 安装前端依赖
cd frontend
npm install

# 安装后端依赖
cd ../backend
pip install -r requirements.txt
```

### 启动开发环境

```bash
# 在根目录运行（会同时启动前端、后端和 Electron）
npm run dev
```

这会启动：
- 前端开发服务器: http://localhost:5173
- Python 后端: http://localhost:8000
- Electron 窗口

### 单独启动

```bash
# 只启动前端
npm run dev:frontend

# 只启动后端
npm run dev:backend

# 只启动 Electron
npm run dev:electron
```

## 构建打包

### 构建前端

```bash
npm run build:frontend
```

### 打包应用

```bash
# Windows 安装包（NSIS + MSI）
npm run build:win

# macOS 安装包
npm run build:mac

# Linux 安装包
npm run build:linux

# 所有平台
npm run dist
```

输出目录：`dist/`

### 测试打包（不生成安装包）

```bash
npm run pack
```

## 功能特性

### Electron 主进程

- 窗口管理
- Python 后端自动启动和管理
- 文件系统访问
- 本地数据存储（electron-store）
- IPC 通信

### 前端功能

- 文件上传（支持原生文件选择器）
- 数据处理和展示
- 设置管理
- 主题切换

### 后端功能

- 文档解析
- 税率查询
- HS 编码优化
- 文件生成

## API 通信

### Electron API

前端可以通过 `window.electronAPI` 访问：

```javascript
// 获取设置
const settings = await window.electronAPI.getSettings();

// 保存设置
await window.electronAPI.saveSettings(settings);

// 选择文件
const files = await window.electronAPI.selectFiles();

// 选择文件夹
const folder = await window.electronAPI.selectFolder();

// 保存文件
const path = await window.electronAPI.saveFile({
  defaultPath: 'output.xlsx',
  filters: [{ name: 'Excel', extensions: ['xlsx'] }]
});

// 打开外部链接
await window.electronAPI.openExternal('https://example.com');

// 获取应用信息
const info = await window.electronAPI.getAppInfo();

// 获取后端状态
const status = await window.electronAPI.getBackendStatus();

// 重启后端
await window.electronAPI.restartBackend();
```

### 环境检测

```javascript
// 检测是否在 Electron 环境
if (window.electronAPI?.isElectron) {
  // Electron 环境
} else {
  // Web 环境
}
```

## 配置说明

### electron-builder.yml

打包配置文件，包含：
- 应用信息（名称、版本、版权）
- 打包目录
- 资源文件
- 平台特定配置（Windows/Mac/Linux）
- 安装包配置

### electron-store

本地数据存储，自动保存在：
- Windows: `%APPDATA%/DutyWise/config.json`
- macOS: `~/Library/Application Support/DutyWise/config.json`
- Linux: `~/.config/DutyWise/config.json`

## 故障排查

### Python 后端启动失败

1. 检查 Python 是否安装：`python --version`
2. 检查依赖是否安装：`pip list`
3. 检查端口是否被占用：`netstat -ano | findstr 8000`
4. 查看 Electron 控制台日志

### 前端无法连接后端

1. 确认后端已启动：访问 http://localhost:8000/docs
2. 检查防火墙设置
3. 检查 CORS 配置

### 打包失败

1. 确认所有依赖已安装
2. 确认前端已构建：`frontend/dist` 目录存在
3. 检查 `build/electron-builder.yml` 配置
4. 查看构建日志

## 更新日志

### v1.0.0 (2026-04-24)

- 初始版本
- Electron 桌面应用框架
- Python 后端集成
- 文件选择和保存
- 本地设置存储

## 许可证

MIT
