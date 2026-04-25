# 快速开始指南

## 1. 环境准备

### 安装 Python
确保已安装 Python 3.8 或更高版本：
```bash
python3 --version
```

### 创建虚拟环境
```bash
cd backend
python3 -m venv venv
source venv/bin/activate  # Linux/Mac
# 或
venv\Scripts\activate  # Windows
```

## 2. 安装依赖

```bash
pip install -r requirements.txt
```

## 3. 配置环境变量

复制环境变量示例文件：
```bash
cp .env.example .env
```

编辑 `.env` 文件，配置必要的参数：
- `DOUBAO_API_KEY`: 豆包 API 密钥（如果使用 AI 功能）
- `REDIS_HOST`: Redis 主机地址（如果使用缓存）
- 其他配置保持默认即可

## 4. 启动服务

### 方式一：使用启动脚本
```bash
./start.sh
```

### 方式二：直接运行
```bash
python -m app.main
```

### 方式三：使用 uvicorn
```bash
uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
```

## 5. 访问 API 文档

启动成功后，访问：
- Swagger UI: http://localhost:8000/docs
- ReDoc: http://localhost:8000/redoc
- 健康检查: http://localhost:8000/health

## 6. 测试 API

### 使用 curl
```bash
# 健康检查
curl http://localhost:8000/health

# 上传文件
curl -X POST http://localhost:8000/api/v1/shipments/upload \
  -F "files=@your_file.xlsx"
```

### 使用测试脚本
```bash
python test_api.py
```

## 7. 项目结构说明

```
backend/
├── app/
│   ├── api/              # API 路由层
│   │   └── routes.py     # 定义所有 API 端点
│   ├── core/             # 核心配置
│   │   └── config.py     # 应用配置和环境变量
│   ├── models/           # 数据模型
│   │   └── schemas.py    # Pydantic 模型定义
│   ├── services/         # 业务逻辑层
│   │   ├── parser.py     # Excel 文件解析
│   │   ├── mapper.py     # 字段映射
│   │   ├── crawler.py    # 税率爬虫
│   │   ├── optimizer.py  # HS编码优化
│   │   └── generator.py  # 文件生成
│   ├── utils/            # 工具函数
│   │   └── helpers.py    # 辅助函数
│   └── main.py           # FastAPI 应用入口
├── requirements.txt      # Python 依赖包
├── .env.example         # 环境变量示例
└── README.md            # 项目说明
```

## 8. 开发说明

### 添加新的 API 端点
在 `app/api/routes.py` 中添加新的路由函数。

### 添加新的服务模块
在 `app/services/` 目录下创建新的服务类。

### 修改配置
在 `app/core/config.py` 中添加新的配置项。

### 添加数据模型
在 `app/models/schemas.py` 中定义新的 Pydantic 模型。

## 9. 常见问题

### 端口被占用
修改 `.env` 文件中的 `PORT` 配置。

### 依赖安装失败
尝试升级 pip：
```bash
pip install --upgrade pip
```

### 模块导入错误
确保在项目根目录运行，并且虚拟环境已激活。

## 10. 下一步

- 准备清关文件模板（template.xlsx）
- 实现税率爬虫的具体逻辑
- 配置 Redis 用于缓存
- 配置 Celery 用于异步任务
- 连接数据库（如需持久化存储）
