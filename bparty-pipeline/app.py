"""
bparty-pipeline FastAPI 入口 (端口 8001)

启动:
    cd bparty-pipeline
    python app.py

端点:
    GET  /health           → 健康检查
    POST /process          → 上传文件 + 模式 → 运行流水线 → 返回 stats + 下载链接
    GET  /download/{name}  → 下载生成的文件
"""

import logging
import shutil
import sys
import uuid
from datetime import datetime
from pathlib import Path

# --- sys.path 注入: 复用 backend/ 代码 ---
# backend 内部使用 `from app.core.config import settings` 等相对导入，
# 需要将 backend/ 目录加入 sys.path 使 `app` 解析为 `backend/app/`。
# _BACKEND_ROOT 必须在 _HERE 之前，否则 bparty-pipeline/app.py 遮蔽 backend/app/ 包。
_HERE = Path(__file__).resolve().parent
_PROJECT_ROOT = _HERE.parent
_BACKEND_ROOT = _PROJECT_ROOT / "backend"

sys.path.insert(0, str(_HERE))           # strategy.*, pipeline (后搜索)
sys.path.insert(0, str(_BACKEND_ROOT))   # app.* → backend/app/* (先搜索)

import yaml
from fastapi import FastAPI, File, Form, UploadFile, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse

from pipeline import _load_config, _resolve_mode_config, run_pipeline
from strategy.bl_parser import BLParser

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger("bpapi")

app = FastAPI(
    title="bparty-pipeline API",
    version="1.0.0",
    description="乙方清关流水线自动化 API",
)

# CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "http://localhost:3000",
        "http://localhost:3001",
        "http://localhost:5173",
        "http://127.0.0.1:3000",
        "http://127.0.0.1:3001",
        "http://127.0.0.1:5173",
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 配置
BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

# 任务存储 (内存)
tasks: dict[str, dict] = {}

# config.yaml 路径
CONFIG_PATH = BASE_DIR / "config.yaml"


@app.get("/health")
async def health():
    return {"status": "ok", "timestamp": datetime.now().isoformat()}


@app.post("/process")
async def process(
    file: UploadFile = File(...),
    mode: str = Form("conservative"),
    use_crawler: bool = Form(True),
    bl_file: UploadFile = File(None),
):
    """
    上传 Excel 文件 + 可选提单PDF → 运行完整流水线 → 返回 stats + download_url

    Args:
        bl_file: 提单PDF文件（可选）。提单上的品名会被保护，在输出中至少出现一行。
    """
    if not file.filename or not file.filename.lower().endswith((".xlsx", ".xls")):
        raise HTTPException(400, "仅支持 .xlsx / .xls 文件")

    # 保存上传文件
    task_id = uuid.uuid4().hex[:12]
    safe_name = f"{task_id}_{file.filename}"
    upload_path = UPLOAD_DIR / safe_name
    with open(upload_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    # 解析提单（如果提供）
    bl_products = []
    if bl_file and bl_file.filename:
        bl_safe_name = f"{task_id}_bl_{bl_file.filename}"
        bl_path = UPLOAD_DIR / bl_safe_name
        with open(bl_path, "wb") as f:
            shutil.copyfileobj(bl_file.file, f)

        try:
            bl_parser = BLParser()
            bl_products = bl_parser.parse_pdf(str(bl_path))
            logger.info(f"提单: {bl_file.filename} → 品名 {bl_products}")
        except Exception as e:
            logger.warning(f"提单解析失败（继续处理）: {e}")
            bl_products = []
        finally:
            try:
                bl_path.unlink()
            except OSError:
                pass

    logger.info(f"任务 {task_id}: 接收文件 {file.filename}, 模式={mode}, 提单品名={bl_products}")

    # 加载配置
    if not CONFIG_PATH.exists():
        raise HTTPException(500, f"配置文件不存在: {CONFIG_PATH}")
    full_cfg = _load_config(str(CONFIG_PATH))
    mode_cfg = _resolve_mode_config(full_cfg, mode)

    try:
        output_path = str(OUTPUT_DIR / f"output_{task_id}.xlsx")
        result = await run_pipeline(
            input_path=str(upload_path),
            output_path=output_path,
            mode_cfg=mode_cfg,
            use_crawler=use_crawler,
            bl_products=bl_products,
        )

        # 记录任务
        output_filename = Path(result["output_file"]).name
        tasks[task_id] = {
            "task_id": task_id,
            "status": "completed",
            "input_file": file.filename,
            "mode": mode,
            "stats": result["stats"],
            "output_filename": output_filename,
            "created_at": datetime.now().isoformat(),
        }

        return JSONResponse({
            "code": 200,
            "message": "流水线执行成功",
            "data": {
                "task_id": task_id,
                "stats": result["stats"],
                "download_url": f"/download/{output_filename}",
            }
        })
    except Exception as e:
        logger.exception(f"任务 {task_id} 失败")
        tasks[task_id] = {
            "task_id": task_id,
            "status": "failed",
            "input_file": file.filename,
            "mode": mode,
            "error": str(e),
            "created_at": datetime.now().isoformat(),
        }
        raise HTTPException(500, f"流水线执行失败: {e}")


@app.get("/download/{filename}")
async def download(filename: str):
    """下载生成的清关文件"""
    file_path = OUTPUT_DIR / filename
    if not file_path.exists():
        raise HTTPException(404, "文件不存在或已被清理")
    return FileResponse(
        path=str(file_path),
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.get("/tasks/{task_id}")
async def get_task(task_id: str):
    """查询任务状态"""
    task = tasks.get(task_id)
    if not task:
        raise HTTPException(404, "任务不存在")
    return {"code": 200, "data": task}


if __name__ == "__main__":
    import uvicorn
    logger.info("启动 bparty-pipeline API 服务 (端口 8001)")
    uvicorn.run(app, host="0.0.0.0", port=8001, log_level="info")
