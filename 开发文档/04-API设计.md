# API 设计

## 4.1 技术栈

- **框架**: FastAPI
- **认证**: JWT
- **文档**: OpenAPI (Swagger)
- **WebSocket**: 实时进度推送

---

## 4.2 API 列表

### 4.2.1 文件上传

**POST /api/v1/shipments/upload**

上传一"票"的文档文件。

**Request**:
```http
POST /api/v1/shipments/upload
Content-Type: multipart/form-data

files: [file1.xlsx, file2.xlsx, file3.xlsx]
```

**Response**:
```json
{
    "code": 200,
    "message": "上传成功",
    "data": {
        "shipment_id": 123,
        "shipment_no": "SHIP20260424001",
        "files": [
            {
                "filename": "04.03装柜信息(1).xlsx",
                "size": 139824,
                "path": "/uploads/20260424/xxx.xlsx"
            }
        ],
        "task_id": "task_abc123"
    }
}
```

---

### 4.2.2 开始处理

**POST /api/v1/shipments/{shipment_id}/process**

开始处理一"票"的数据。

**Request**:
```json
{
    "options": {
        "exclude_anti_dumping": false,  // 是否排除有反倾销的
        "min_similarity": 0.6            // 最低相似度阈值
    }
}
```

**Response**:
```json
{
    "code": 200,
    "message": "处理已开始",
    "data": {
        "task_id": "task_abc123",
        "estimated_time": 120  // 预计耗时（秒）
    }
}
```

---

### 4.2.3 查询处理状态

**GET /api/v1/tasks/{task_id}/status**

查询任务处理状态。

**Response**:
```json
{
    "code": 200,
    "data": {
        "task_id": "task_abc123",
        "status": "processing",  // processing/completed/failed
        "progress": 65,          // 0-100
        "current_step": "优化HS编码",
        "steps": [
            {"name": "解析文档", "status": "completed"},
            {"name": "字段映射", "status": "completed"},
            {"name": "查询税率", "status": "completed"},
            {"name": "优化HS编码", "status": "processing"},
            {"name": "生成文件", "status": "pending"}
        ],
        "error_message": null
    }
}
```

---

### 4.2.4 获取处理结果

**GET /api/v1/shipments/{shipment_id}/result**

获取处理完成后的结果。

**Response**:
```json
{
    "code": 200,
    "data": {
        "shipment_id": 123,
        "shipment_no": "SHIP20260424001",
        "status": "completed",
        "statistics": {
            "total_items": 439,
            "optimized_items": 312,
            "optimization_rate": 71.1,
            "estimated_tax_saved": 15000
        },
        "files": {
            "output_file": "/downloads/清关_20260424_120000.xlsx",
            "log_file": "/downloads/日志_20260424_120000.xlsx"
        },
        "warnings": [
            {
                "type": "high_risk",
                "count": 5,
                "message": "有5个商品存在高风险"
            },
            {
                "type": "conflict",
                "count": 2,
                "message": "有2个字段存在数据冲突"
            }
        ]
    }
}
```

---

### 4.2.5 下载文件

**GET /api/v1/files/download/{file_id}**

下载生成的文件。

**Response**:
```http
Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
Content-Disposition: attachment; filename="清关_20260424_120000.xlsx"

[文件二进制数据]
```

---

### 4.2.6 获取票据列表

**GET /api/v1/shipments**

获取用户的票据列表。

**Query Parameters**:
- `page`: 页码（默认1）
- `page_size`: 每页数量（默认20）
- `status`: 状态筛选（可选）

**Response**:
```json
{
    "code": 200,
    "data": {
        "total": 156,
        "page": 1,
        "page_size": 20,
        "items": [
            {
                "shipment_id": 123,
                "shipment_no": "SHIP20260424001",
                "status": "completed",
                "total_items": 439,
                "optimized_items": 312,
                "created_at": "2026-04-24T10:00:00",
                "completed_at": "2026-04-24T10:05:30"
            },
            ...
        ]
    }
}
```

---

### 4.2.7 获取票据详情

**GET /api/v1/shipments/{shipment_id}**

获取票据的详细信息。

**Response**:
```json
{
    "code": 200,
    "data": {
        "shipment_id": 123,
        "shipment_no": "SHIP20260424001",
        "status": "completed",
        "created_at": "2026-04-24T10:00:00",
        "completed_at": "2026-04-24T10:05:30",
        "input_files": [
            {"filename": "04.03装柜信息(1).xlsx", "size": 139824}
        ],
        "statistics": {
            "total_items": 439,
            "optimized_items": 312,
            "high_risk_items": 5,
            "conflict_items": 2
        },
        "items": [
            {
                "item_id": 1001,
                "original_name_cn": "阀门配件",
                "original_hs_code": "8481901000",
                "original_tax_rate": "25%",
                "optimized_hs_code": "8481900085",
                "optimized_tax_rate": "Free",
                "is_optimized": true,
                "similarity_score": 0.85,
                "risk_level": "MEDIUM",
                "risk_warnings": ["⚠️ 有反倾销标记"]
            },
            ...
        ]
    }
}
```

---

### 4.2.8 获取优化日志

**GET /api/v1/shipments/{shipment_id}/logs**

获取优化日志详情。

**Query Parameters**:
- `log_type`: 日志类型（optimization/conflict/mapping）

**Response**:
```json
{
    "code": 200,
    "data": {
        "optimization_logs": [
            {
                "item_id": 1001,
                "original_name": "阀门配件",
                "original_hs_code": "8481901000",
                "original_tax_rate": "25%",
                "new_hs_code": "8481900085",
                "new_tax_rate": "Free",
                "reason": "税率更低",
                "similarity_score": 0.85,
                "risk_warning": "⚠️ 有反倾销标记"
            },
            ...
        ],
        "conflict_logs": [
            {
                "field": "总重量",
                "values": ["15650kg", "15800kg"],
                "resolution": "已空着，请手动填写"
            },
            ...
        ],
        "mapping_logs": [
            {
                "source_field": "品名",
                "target_field": "中文品名",
                "confidence": 0.95,
                "method": "keyword_match"
            },
            ...
        ]
    }
}
```

---

## 4.3 WebSocket API

### 4.3.1 连接

**WS /api/v1/ws/tasks/{task_id}**

实时推送任务进度。

**消息格式**:
```json
{
    "type": "progress",
    "data": {
        "progress": 65,
        "current_step": "优化HS编码",
        "message": "正在处理第 285/439 个商品"
    }
}
```

**消息类型**:
- `progress`: 进度更新
- `completed`: 任务完成
- `error`: 任务失败

---

## 4.4 错误码

| 错误码 | 说明 |
|-------|------|
| 200 | 成功 |
| 400 | 请求参数错误 |
| 401 | 未授权 |
| 404 | 资源不存在 |
| 500 | 服务器内部错误 |
| 1001 | 文件格式不支持 |
| 1002 | 文件解析失败 |
| 1003 | 字段映射失败 |
| 1004 | 税率查询失败 |
| 1005 | HS编码优化失败 |
| 1006 | 文件生成失败 |

---

## 4.5 API 实现示例

### 4.5.1 上传文件

```python
from fastapi import FastAPI, UploadFile, File
from typing import List

app = FastAPI()

@app.post("/api/v1/shipments/upload")
async def upload_files(files: List[UploadFile] = File(...)):
    """上传文件"""
    # 1. 创建票据记录
    shipment = create_shipment()

    # 2. 保存文件
    saved_files = []
    for file in files:
        file_path = save_uploaded_file(file, shipment.id)
        saved_files.append({
            "filename": file.filename,
            "size": file.size,
            "path": file_path
        })

    # 3. 创建异步任务
    task_id = create_processing_task(shipment.id, saved_files)

    return {
        "code": 200,
        "message": "上传成功",
        "data": {
            "shipment_id": shipment.id,
            "shipment_no": shipment.shipment_no,
            "files": saved_files,
            "task_id": task_id
        }
    }
```

### 4.5.2 开始处理

```python
from celery import Celery

celery_app = Celery('customs', broker='redis://localhost:6379/0')

@app.post("/api/v1/shipments/{shipment_id}/process")
async def process_shipment(shipment_id: int, options: dict):
    """开始处理"""
    # 1. 验证票据存在
    shipment = get_shipment(shipment_id)
    if not shipment:
        return {"code": 404, "message": "票据不存在"}

    # 2. 创建Celery任务
    task = process_shipment_task.delay(shipment_id, options)

    return {
        "code": 200,
        "message": "处理已开始",
        "data": {
            "task_id": task.id,
            "estimated_time": 120
        }
    }

@celery_app.task
def process_shipment_task(shipment_id: int, options: dict):
    """处理票据（Celery任务）"""
    try:
        # 1. 解析文档
        update_task_progress(task.id, 10, "解析文档")
        parsed_data = DocumentParser().parse_files(shipment.input_files)

        # 2. 字段映射
        update_task_progress(task.id, 30, "字段映射")
        mapped_data = FieldMapper().map_fields(parsed_data)

        # 3. 查询税率
        update_task_progress(task.id, 50, "查询税率")
        tax_data = TaxRateCrawler().batch_search(mapped_data)

        # 4. 优化HS编码
        update_task_progress(task.id, 70, "优化HS编码")
        optimized_data = HSCodeOptimizer().optimize_batch(mapped_data, tax_data)

        # 5. 生成文件
        update_task_progress(task.id, 90, "生成文件")
        output_file = FileGenerator().generate(optimized_data)
        log_file = Logger().create_log_file(optimized_data)

        # 6. 更新数据库
        update_shipment(shipment_id, {
            "status": "completed",
            "output_file": output_file,
            "log_file": log_file
        })

        update_task_progress(task.id, 100, "完成")

    except Exception as e:
        update_shipment(shipment_id, {"status": "failed"})
        update_task_progress(task.id, -1, f"失败: {str(e)}")
        raise
```

### 4.5.3 WebSocket 推送

```python
from fastapi import WebSocket

@app.websocket("/api/v1/ws/tasks/{task_id}")
async def websocket_endpoint(websocket: WebSocket, task_id: str):
    """WebSocket连接"""
    await websocket.accept()

    try:
        while True:
            # 从Redis获取任务状态
            task_status = get_task_status(task_id)

            # 推送给客户端
            await websocket.send_json({
                "type": "progress",
                "data": task_status
            })

            # 如果任务完成，断开连接
            if task_status["status"] in ["completed", "failed"]:
                break

            await asyncio.sleep(1)

    except WebSocketDisconnect:
        pass
```

---

## 4.6 API 测试

### 4.6.1 使用 curl

```bash
# 上传文件
curl -X POST http://localhost:8000/api/v1/shipments/upload \
  -F "files=@04.03装柜信息(1).xlsx" \
  -F "files=@ZIMUSNH21275749--04.21.xls"

# 开始处理
curl -X POST http://localhost:8000/api/v1/shipments/123/process \
  -H "Content-Type: application/json" \
  -d '{"options": {"exclude_anti_dumping": false}}'

# 查询状态
curl http://localhost:8000/api/v1/tasks/task_abc123/status

# 下载文件
curl -O http://localhost:8000/api/v1/files/download/file_123
```

### 4.6.2 使用 Python

```python
import requests

# 上传文件
files = [
    ('files', open('04.03装柜信息(1).xlsx', 'rb')),
    ('files', open('ZIMUSNH21275749--04.21.xls', 'rb'))
]
response = requests.post('http://localhost:8000/api/v1/shipments/upload', files=files)
data = response.json()

shipment_id = data['data']['shipment_id']
task_id = data['data']['task_id']

# 查询状态
while True:
    response = requests.get(f'http://localhost:8000/api/v1/tasks/{task_id}/status')
    status = response.json()['data']

    print(f"进度: {status['progress']}% - {status['current_step']}")

    if status['status'] in ['completed', 'failed']:
        break

    time.sleep(2)

# 下载结果
response = requests.get(f'http://localhost:8000/api/v1/shipments/{shipment_id}/result')
result = response.json()['data']

output_file = result['files']['output_file']
requests.get(f'http://localhost:8000{output_file}', stream=True)
```
