# DutyWise 未完成任务清单

> 最后更新: 2026-04-26

## 1. 爬虫API登录对接 (codeflagai.com) ✅ 已完成

**状态**: AES-128-ECB加密登录 + CusAuthorization JWT + classification/search 全部打通

**实现细节**:
- 登录: `POST /xhqUser/login` — body用AES-128-ECB(key="imageBatchCompon")加密
- Token: 从响应头 `CusAuthorization` 提取JWT
- 搜索: `POST /classification/search` — 带 CusAuthorization header
- 税率字段: `classificationCodeList[].importTariffRate`
- 反倾销: `classificationCodeList[].antiDumpingCountervailingRate`
- 每次batch前强制重新登录，遇1401自动清除token并重试

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

## 3. ~~前端硬编码 localhost~~ → 已降级 🟢 低优先级

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
