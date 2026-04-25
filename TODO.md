# DutyWise 未完成任务清单

> 最后更新: 2026-04-25

## 1. 爬虫API登录对接 (codeflagai.com) 🔴 高优先级

**当前状态**: 登录始终返回 `1401 登录过期`，所有HS编码查询回退到内置税率库

**根因**: 
- 正确的API Host是 `api.codeflagai.com`，而非当前配置的 `www.codeflagai.com`
- `api.codeflagai.com` 要求 `appId` 参数 (返回码510: "缺少必要参数appId")
- 当前未知: 正确的 `appId` 值、登录接口完整参数、Token获取方式

**待调查**:
- [ ] 通过浏览器抓包 codeflagai.com 登录流程，获取完整API地址和参数
- [ ] 找到 `appId` 或 `client_id` 值
- [ ] 确认 Token 存储位置 (header/cookie/response body)
- [ ] 更新 `.env` 中的 CRAWLER_BASE_URL 和爬虫登录逻辑

**相关文件**:
- `backend/app/services/crawler.py` — `_do_login()`, `_api_search()`
- `backend/.env` — `CRAWLER_BASE_URL`, `CRAWLER_USERNAME`, `CRAWLER_PASSWORD`

---

## 2. Electron 二进制下载 🟡 中优先级

**当前状态**: Electron 二进制因网络问题下载失败，目前仅测试了 Web 模式 (`npm run dev`)

**待处理**:
- [ ] 配置 Electron 镜像源 (可选)
- [ ] 重新执行 Electron 安装
- [ ] 测试桌面端完整流程

---

## 3. 前端硬编码 localhost 🟡 中优先级

**当前状态**: `ResultCard.tsx` 中的下载链接硬编码为 `http://localhost:8000`

```ts
// frontend/src/components/ResultCard.tsx:19
window.open(`http://localhost:8000/api/v1/files/download?path=...`)
```

**建议**:
- [ ] 使用 Vite 代理或环境变量统一管理 API 地址
- [ ] 或使用相对路径 `/api/v1/files/download?path=...`（配合 Vite proxy）

---

## 4. 测试覆盖 🟢 低优先级

**当前状态**: 仅有手动端到端测试，无自动化测试

**建议**:
- [ ] 添加后端 API 单元测试 (pytest)
- [ ] 添加前端组件测试
- [ ] 添加 FieldMapper 和 Optimizer 的核心逻辑测试

---

## 5. 内置税率库扩展 🟢 低优先级

**当前状态**: `RATE_DB` 仅覆盖 8517/8518/8471/7326/3926/6702 六类HS编码

**建议**:
- [ ] 逐步补充更多常用HS编码的税率数据
- [ ] 考虑将数据库从代码中移到独立配置文件或数据库

---

## 已完成修复 ✅

| 问题 | 修复内容 |
|------|---------|
| FieldMapper 返回0条 | `"品名" in "英文品名"` 误匹配 → 评分制精确匹配 |
| Optimizer 返回0条 | 阈值 `0.7→0.55`，修复 `Free` 税率解析 |
| 下载按钮无响应 | 创建 template.xlsx + 添加下载端点 |
| 前端 task_id 错乱 | 使用 process 返回的 task_id 而非 upload 的 |
| TypeScript 编译错误 | 修复未使用导入、隐式any、vite类型声明 |
| Tailwind CSS 无效 | 全部改为 inline styles |
| Pydantic 验证报错 | `extra="ignore"` 兼容 .env 多余字段 |
| 爬虫登录缓存问题 | 每次 batch 前强制重登 + 1401 自动重试 |
