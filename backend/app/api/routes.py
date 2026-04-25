"""
API 路由
"""
from fastapi import APIRouter, UploadFile, File, HTTPException, BackgroundTasks
from typing import List
import shutil
from pathlib import Path
import logging

from app.models.schemas import (
    APIResponse, ProcessOptions, TaskProgress,
    TaskStatus, UploadFileInfo
)
from app.core.config import settings
from app.utils.helpers import (
    generate_shipment_no, generate_task_id,
    ensure_dir, is_allowed_file
)
from app.services.parser import DocumentParser
from app.services.mapper import FieldMapper
from app.services.crawler import TaxRateCrawler
from app.services.optimizer import HSCodeOptimizer
from app.services.generator import FileGenerator

router = APIRouter(prefix="/api/v1", tags=["shipments"])
logger = logging.getLogger(__name__)

# 简单的内存存储（生产环境应使用数据库）
shipments_db = {}
tasks_db = {}


@router.post("/shipments/upload")
async def upload_files(files: List[UploadFile] = File(...)):
    """上传文件"""
    try:
        # 验证文件
        for file in files:
            if not is_allowed_file(file.filename, settings.ALLOWED_EXTENSIONS):
                raise HTTPException(
                    status_code=400,
                    detail=f"不支持的文件格式: {file.filename}"
                )

        # 创建票据
        shipment_id = len(shipments_db) + 1
        shipment_no = generate_shipment_no()

        # 保存文件
        upload_dir = ensure_dir(settings.UPLOAD_DIR) / shipment_no
        upload_dir.mkdir(exist_ok=True)

        saved_files = []
        file_paths = []

        for file in files:
            file_path = upload_dir / file.filename

            with file_path.open("wb") as buffer:
                shutil.copyfileobj(file.file, buffer)

            saved_files.append({
                "filename": file.filename,
                "size": file_path.stat().st_size,
                "path": str(file_path)
            })
            file_paths.append(str(file_path))

        # 保存票据信息
        shipments_db[shipment_id] = {
            "shipment_id": shipment_id,
            "shipment_no": shipment_no,
            "status": "uploaded",
            "files": saved_files,
            "file_paths": file_paths,
            "created_at": None
        }

        task_id = generate_task_id()

        return APIResponse(
            code=200,
            message="上传成功",
            data={
                "shipment_id": shipment_id,
                "shipment_no": shipment_no,
                "files": saved_files,
                "task_id": task_id
            }
        )

    except Exception as e:
        logger.error(f"上传文件失败: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))


@router.post("/shipments/{shipment_id}/process")
async def process_shipment(
    shipment_id: int,
    options: ProcessOptions,
    background_tasks: BackgroundTasks
):
    """开始处理票据"""
    try:
        # 验证票据存在
        shipment = shipments_db.get(shipment_id)
        if not shipment:
            raise HTTPException(status_code=404, detail="票据不存在")

        # 创建任务
        task_id = generate_task_id()

        tasks_db[task_id] = {
            "task_id": task_id,
            "shipment_id": shipment_id,
            "status": TaskStatus.PENDING,
            "progress": 0,
            "current_step": "等待开始",
            "error_message": None
        }

        # 后台处理
        background_tasks.add_task(
            process_shipment_task,
            task_id,
            shipment_id,
            options.dict()
        )

        return APIResponse(
            code=200,
            message="处理已开始",
            data={
                "task_id": task_id,
                "estimated_time": 120
            }
        )

    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"开始处理失败: {str(e)}")
        raise HTTPException(status_code=500, detail=str(e))


@router.get("/tasks/{task_id}/status")
async def get_task_status(task_id: str):
    """查询任务状态"""
    task = tasks_db.get(task_id)

    if not task:
        raise HTTPException(status_code=404, detail="任务不存在")

    return APIResponse(
        code=200,
        data=task
    )


@router.get("/shipments/{shipment_id}/result")
async def get_shipment_result(shipment_id: int):
    """获取处理结果"""
    shipment = shipments_db.get(shipment_id)

    if not shipment:
        raise HTTPException(status_code=404, detail="票据不存在")

    if shipment["status"] != "completed":
        raise HTTPException(status_code=400, detail="处理尚未完成")

    return APIResponse(
        code=200,
        data=shipment.get("result", {})
    )


async def process_shipment_task(task_id: str, shipment_id: int, options: dict):
    """处理票据任务（后台任务）"""
    try:
        shipment = shipments_db[shipment_id]
        file_paths = shipment["file_paths"]

        # 更新任务状态
        def update_progress(progress: int, step: str):
            tasks_db[task_id].update({
                "status": TaskStatus.PROCESSING,
                "progress": progress,
                "current_step": step
            })

        # 1. 解析文档
        update_progress(10, "解析文档")
        parser = DocumentParser()
        parsed_data = parser.parse_files(file_paths)

        # 2. 字段映射
        update_progress(30, "字段映射")
        mapper = FieldMapper()
        mapped_result = mapper.map_fields(parsed_data)
        mapped_data = mapped_result["items"]

        # 3. 查询税率
        update_progress(50, "查询税率")
        crawler = TaxRateCrawler()
        tax_data = await crawler.batch_search(mapped_data)

        # 4. 优化HS编码
        update_progress(70, "优化HS编码")
        optimizer = HSCodeOptimizer(options)
        optimized_result = optimizer.optimize_batch(mapped_data, tax_data)

        # 5. 生成文件
        update_progress(90, "生成文件")
        output_dir = ensure_dir(settings.OUTPUT_DIR)

        # 这里需要模板文件路径（实际使用时需要配置）
        template_path = "template.xlsx"  # 需要实际的模板文件

        generator = FileGenerator()

        # 如果模板存在，生成文件
        if Path(template_path).exists():
            output_file = generator.generate(
                optimized_result["items"],
                template_path,
                str(output_dir)
            )

            log_file = generator.generate_log_file(
                {
                    "optimization_logs": optimized_result["optimization_logs"],
                    "conflicts": [],
                    "mapping_logs": mapped_result.get("mapping_logs", [])
                },
                str(output_dir)
            )
        else:
            output_file = None
            log_file = None

        # 更新票据状态
        shipments_db[shipment_id].update({
            "status": "completed",
            "result": {
                "shipment_id": shipment_id,
                "shipment_no": shipment["shipment_no"],
                "status": "completed",
                "statistics": {
                    "total_items": len(mapped_data),
                    "optimized_items": len(optimized_result["optimization_logs"]),
                    "optimization_rate": len(optimized_result["optimization_logs"]) / len(mapped_data) * 100 if mapped_data else 0
                },
                "files": {
                    "output_file": output_file,
                    "log_file": log_file
                }
            }
        })

        # 更新任务状态
        tasks_db[task_id].update({
            "status": TaskStatus.COMPLETED,
            "progress": 100,
            "current_step": "完成"
        })

    except Exception as e:
        logger.error(f"处理任务失败: {str(e)}")

        tasks_db[task_id].update({
            "status": TaskStatus.FAILED,
            "error_message": str(e)
        })

        shipments_db[shipment_id]["status"] = "failed"
