"""
API 路由
"""
from fastapi import APIRouter, UploadFile, File, HTTPException
from fastapi.responses import FileResponse
from typing import List
import shutil
import asyncio
from pathlib import Path
import logging
from datetime import datetime

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
            "created_at": datetime.now().isoformat()
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
    options: ProcessOptions
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
            "errorMessage": None
        }

        # 后台处理（用asyncio.create_task在主event loop中运行）
        asyncio.create_task(process_shipment_task_async(task_id, shipment_id, options.dict()))

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
        message="查询成功",
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
        message="查询成功",
        data=shipment.get("result", {})
    )


@router.get("/files/download")
async def download_file(path: str, name: str = "output.xlsx"):
    """下载生成的文件"""
    file_path = Path(path)
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="文件不存在")
    return FileResponse(
        path=str(file_path),
        filename=name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@router.get("/shipments")
async def list_shipments(page: int = 1, page_size: int = 20):
    """列出所有票据（历史记录）"""
    all_shipments = []
    for sid, s in shipments_db.items():
        result_data = s.get("result", {}) or {}
        item = {
            "id": s["shipment_id"],
            "shipmentNo": s["shipment_no"],
            "status": s.get("status", "unknown"),
            "totalItems": result_data.get("statistics", {}).get("total_items", 0),
            "optimizedItems": result_data.get("statistics", {}).get("optimized_items", 0),
            "createdAt": s.get("created_at"),
            "completedAt": None,
        }
        if s.get("status") == "completed":
            item["completedAt"] = result_data.get("completed_at")
        all_shipments.append(item)

    all_shipments.sort(key=lambda x: x["id"], reverse=True)
    total = len(all_shipments)
    start = (page - 1) * page_size
    paged = all_shipments[start:start + page_size]

    return APIResponse(
        code=200,
        message="查询成功",
        data={"items": paged, "total": total}
    )


async def process_shipment_task_async(task_id: str, shipment_id: int, options: dict):
    """处理票据任务（异步，运行在主event loop中）"""
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

        # 检查是否包含产品报关信息
        has_product_data = False
        for file_data in parsed_data["files"]:
            for sheet in file_data["sheets"]:
                if sheet.get("has_product_fields", True):
                    has_product_data = True
                    break
            if has_product_data:
                break

        if not has_product_data:
            # 所有sheet都不包含产品报关字段（品名/HS编码/申报价值等）
            error_msg = (
                "上传的文件不包含产品报关信息（品名、HS编码、申报价值等字段）。"
                "请确认上传的是产品报关明细表，而非物流装箱清单。"
            )
            logger.warning(f"任务 {task_id}: {error_msg}")
            tasks_db[task_id].update({
                "status": TaskStatus.FAILED,
                "errorMessage": error_msg,
                "progress": 100,
                "current_step": "文件格式不支持"
            })
            shipments_db[shipment_id]["status"] = "failed"
            return

        # 2. 字段映射
        update_progress(30, "字段映射")
        mapper = FieldMapper()
        mapped_result = mapper.map_fields(parsed_data)
        mapped_data = mapped_result["items"]

        if not mapped_data:
            error_msg = (
                "未能从文件中识别出任何商品信息。"
                "请确认文件包含中文品名、HS编码等必要字段，"
                "且数据格式与系统模板一致。"
            )
            logger.warning(f"任务 {task_id}: {error_msg}")
            tasks_db[task_id].update({
                "status": TaskStatus.FAILED,
                "errorMessage": error_msg,
                "progress": 100,
                "current_step": "字段映射失败"
            })
            shipments_db[shipment_id]["status"] = "failed"
            return

        # 3. 查询税率（直接await，运行在主event loop中）
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

        # 模板文件路径
        template_path = settings.UPLOAD_DIR + "/../template.xlsx"

        generator = FileGenerator()
        output_file = None
        log_file = None

        # 始终生成日志文件（不需要模板）
        try:
            log_file = generator.generate_log_file(
                {
                    "optimization_logs": optimized_result["optimization_logs"],
                    "conflicts": [],
                    "mapping_logs": mapped_result.get("mapping_logs", [])
                },
                str(output_dir)
            )
        except Exception as e:
            logger.warning(f"生成日志文件失败: {e}")

        # 如果有模板，生成清关文件（使用全部映射数据）
        if Path(template_path).exists():
            try:
                output_file = generator.generate(
                    optimized_result["items"],
                    template_path,
                    str(output_dir),
                    optimization_logs=optimized_result.get("optimization_logs", [])
                )
            except Exception as e:
                logger.warning(f"生成清关文件失败: {e}")
        else:
            logger.warning(f"模板文件不存在: {template_path}")

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
            "errorMessage": str(e)
        })

        shipments_db[shipment_id]["status"] = "failed"
