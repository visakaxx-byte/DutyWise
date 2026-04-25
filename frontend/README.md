# 清关优化系统 - 前端

## 快速开始

### 1. 安装依赖

```bash
cd frontend
npm install
```

### 2. 启动开发服务器

```bash
npm run dev
```

访问：http://localhost:3000

### 3. 构建生产版本

```bash
npm run build
```

---

## 项目结构

```
frontend/
├── src/
│   ├── components/          # 组件
│   │   ├── FileUpload.tsx   # 文件上传
│   │   ├── ProgressBar.tsx  # 进度条
│   │   ├── ResultCard.tsx   # 结果卡片
│   │   └── Layout.tsx       # 布局
│   ├── pages/               # 页面
│   │   ├── Home.tsx         # 首页
│   │   ├── History.tsx      # 历史记录
│   │   ├── Settings.tsx     # 设置
│   │   └── Detail.tsx       # 详情
│   ├── store/               # 状态管理
│   │   ├── useAppStore.ts   # 全局状态
│   │   └── useSettingsStore.ts  # 设置状态
│   ├── services/            # API服务
│   │   └── api.ts
│   ├── types/               # 类型定义
│   │   └── index.ts
│   ├── App.tsx
│   └── main.tsx
├── package.json
├── vite.config.ts
└── tsconfig.json
```

---

## 功能说明

### 首页
- 拖拽上传文件（支持 .xlsx, .xls）
- 实时显示处理进度
- 显示处理结果和统计
- 下载生成的文件

### 设置页
- 配置 LLM（Base URL、API Key、Model ID）
- 预设模型快速选择（豆包）
- 测试连接功能
- 优化参数设置
- 爬虫账号配置

### 历史记录页
- 查看所有处理记录
- 搜索和筛选
- 查看详情
- 重新下载文件

---

## 环境变量

创建 `.env` 文件：

```bash
VITE_API_BASE_URL=http://localhost:8000/api/v1
```

---

## 技术栈

- **React 18** - UI框架
- **TypeScript** - 类型安全
- **Ant Design** - UI组件库
- **Zustand** - 状态管理
- **Axios** - HTTP客户端
- **React Dropzone** - 文件上传
- **Vite** - 构建工具

---

## 开发指南

### 添加新页面

1. 在 `src/pages/` 创建新组件
2. 在 `src/App.tsx` 添加路由
3. 在侧边栏添加导航链接

### 添加新API

1. 在 `src/services/api.ts` 添加API函数
2. 在组件中调用

### 状态管理

使用 Zustand 管理全局状态：

```typescript
import { useSettingsStore } from './store/useSettingsStore';

const { settings, updateSettings } = useSettingsStore();
```

---

## 常见问题

### Q: 如何修改API地址？

修改 `.env` 文件中的 `VITE_API_BASE_URL`

### Q: 如何添加新的预设模型？

在 `src/pages/Settings.tsx` 的 `presetModels` 数组中添加

### Q: 如何自定义主题？

修改 `src/App.tsx` 中的 Ant Design ConfigProvider

---

## 部署

### Docker 部署

```bash
docker build -t customs-frontend .
docker run -p 3000:80 customs-frontend
```

### Nginx 部署

```bash
npm run build
cp -r dist/* /var/www/html/
```

---

## 许可证

MIT
