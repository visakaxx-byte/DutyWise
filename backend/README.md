# FastAPI 清关文件处理系统

清关文件处理系统的后端服务，基于 FastAPI 构建。

## 功能特性

- 文档解析：支持 Excel 文件解析
- 字段映射：智能映射到标准字段
- 税率查询：爬取税率信息
- HS编码优化：自动优化选择最优编码
- 文件生成：生成清关文件和日志

## 项目结构

```
backend/
├── app/
│   ├── api/              # API 路由
│   │   └── routes.py
│   ├── core/             # 核心配置
│   │   └── config.py
│   ├── models/           # 数据模型
│   │   └── schemas.py
│   ├── services/         # 业务逻辑
│   │   ├── parser.py     # 文档解析
│   │   ├── mapper.py     # 字段映射
│   │   ├── crawler.py    # 税率爬虫
│   │   ├── optimizer.py  # HS编码优化
│   │   └── generator.py  # 文件生成
│   ├── utils/            # 工具函数
│   │   └── helpers.py
│   └── main.py           # 应用入口
├── requirements.txt      # 依赖包
└── .env.example         # 环境变量示例

## 快速开始

### 1. 安装依赖

```bash
pip install -r requirements.txt
```

### 2. 配置环境变量

```bash
cp .env.example .env
# 编辑 .env 文件，填入实际配置
```

### 3. 启动服务

```bash
# 开发模式
python -m app.main

# 或使用 uvicorn
uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
```

### 4. 访问 API 文档

- Swagger UI: http://localhost:8000/docs
- ReDoc: http://localhost:8000/redoc

## API 接口

### 上传文件
```
POST /api/v1/shipments/upload
```

### 开始处理
```
POST /api/v1/shipments/{shipment_id}/process
```

### 查询任务状态
```
GET /api/v1/tasks/{task_id}/status
```

### 获取处理结果
```
GET /api/v1/shipments/{shipment_id}/result
```

## 开发说明

### 核心模块

1. **DocumentParser**: 解析 Excel 文件
2. **FieldMapper**: 字段映射
3. **TaxRateCrawler**: 税率查询（需要实现实际爬虫逻辑）
4. **HSCodeOptimizer**: HS编码优化
5. **FileGenerator**: 生成清关文件和日志

### 注意事项

- 税率爬虫模块需要根据实际网站结构实现
- 生产环境建议使用数据库替代内存存储
- 需要准备清关文件模板（template.xlsx）
- 建议配置 Redis 用于缓存
- 可选配置 Celery 用于异步任务处理

## 依赖服务

- Redis（可选）：用于缓存和任务队列
- Celery（可选）：用于异步任务处理

## 许可证

MIT
