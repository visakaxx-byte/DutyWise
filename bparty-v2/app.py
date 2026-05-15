from __future__ import annotations

import asyncio
import json
import logging
import shutil
import uuid
from datetime import datetime
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from engine import build_clearance, build_precheck


BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
STATIC_DIR = BASE_DIR / "static"
RUNTIME_DIR = BASE_DIR / "runtime"
PRECHECK_DIR = RUNTIME_DIR / "prechecks"
TASK_DIR = RUNTIME_DIR / "tasks"
logger = logging.getLogger("bparty-v2")

UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
PRECHECK_DIR.mkdir(parents=True, exist_ok=True)
TASK_DIR.mkdir(parents=True, exist_ok=True)
RUNNING_TASKS: dict[str, asyncio.Task] = {}

app = FastAPI(title="BParty V2", version="0.1.0")
app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "http://localhost:8002",
        "http://127.0.0.1:8002",
        "http://localhost:8003",
        "http://127.0.0.1:8003",
        "http://localhost:8004",
        "http://127.0.0.1:8004",
        "http://localhost:8005",
        "http://127.0.0.1:8005",
        "http://localhost:8006",
        "http://127.0.0.1:8006",
        "http://localhost:8007",
        "http://127.0.0.1:8007",
        "http://localhost:8008",
        "http://127.0.0.1:8008",
        "http://localhost:8012",
        "http://127.0.0.1:8012",
        "http://localhost:3000",
        "http://localhost:3001",
        "http://127.0.0.1:3000",
        "http://127.0.0.1:3001",
        "null",
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


def save_upload(upload: UploadFile, prefix: str) -> Path:
    if not upload or not upload.filename:
        raise HTTPException(status_code=400, detail=f"缺少{prefix}文件")
    safe_name = Path(upload.filename).name
    target = UPLOAD_DIR / f"{uuid.uuid4().hex[:10]}_{safe_name}"
    with open(target, "wb") as handle:
        shutil.copyfileobj(upload.file, handle)
    return target


def pick_upload(*uploads: Optional[UploadFile]) -> Optional[UploadFile]:
    for upload in uploads:
        if upload and upload.filename:
            return upload
    return None


def now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


def write_json(path: Path, payload: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_suffix(path.suffix + ".tmp")
    temp.write_text(json.dumps(payload, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    temp.replace(path)


def read_json(path: Path) -> dict:
    if not path.exists():
        raise HTTPException(status_code=404, detail="记录不存在或已过期")
    return json.loads(path.read_text(encoding="utf-8"))


def precheck_path(precheck_id: str) -> Path:
    return PRECHECK_DIR / f"{Path(precheck_id).name}.json"


def task_path(task_id: str) -> Path:
    return TASK_DIR / f"{Path(task_id).name}.json"


def public_precheck_record(record: dict) -> dict:
    return {
        "precheck_id": record["precheck_id"],
        "precheck": record.get("precheck", {}),
        "manifest": record.get("manifest", {}),
        "bill": record.get("bill", {}),
    }


def public_task_record(record: dict) -> dict:
    payload = dict(record)
    if payload.get("output_file"):
        payload["download_url"] = f"/api/download/{Path(payload['output_file']).name}"
    payload.pop("manifest_path", None)
    payload.pop("bill_path", None)
    return payload


def update_task(task_id: str, **updates) -> dict:
    path = task_path(task_id)
    try:
        record = read_json(path)
    except HTTPException:
        record = {"task_id": task_id}
    record.update(updates)
    record["updated_at"] = now_iso()
    write_json(path, record)
    return record


async def run_clearance_job(task_id: str, precheck_record: dict) -> None:
    input_data = precheck_record["input"]
    query_cache = precheck_record.get("query_cache") or {"product": {}, "hs": {}, "bill": {}}
    update_task(
        task_id,
        status="running",
        phase="queued",
        progress=1,
        message="任务已启动",
        started_at=now_iso(),
        manifest_path=input_data["manifest_path"],
        bill_path=input_data["bill_path"],
    )

    async def progress_callback(payload: dict) -> None:
        update_task(
            task_id,
            status="running",
            phase=payload.get("stage") or "running",
            progress=payload.get("progress", 0),
            message=payload.get("message") or "",
            progress_detail=payload,
        )

    try:
        result = await build_clearance(
            input_data["manifest_path"],
            input_data["bill_path"],
            OUTPUT_DIR,
            input_data.get("requested_profile", "auto"),
            target_tax_amount=input_data["target_tax_amount"],
            target_item_count=input_data["target_item_count"],
            query_cache=query_cache,
            progress_callback=progress_callback,
        )
        update_task(
            task_id,
            status="succeeded",
            phase="done",
            progress=100,
            message="处理完成",
            completed_at=now_iso(),
            stats=result["stats"],
            flow=result["flow"],
            manifest=result["manifest"],
            bill=result["bill"],
            bill_categories=result["bill_categories"],
            output_rows=result["output_rows"],
            filter_summary=result.get("filter_summary", {}),
            output_file=result["output_file"],
        )
    except Exception as exc:
        logger.exception("job failed task_id=%s", task_id)
        update_task(
            task_id,
            status="failed",
            phase="failed",
            progress=100,
            message=str(exc),
            error=str(exc),
            completed_at=now_iso(),
        )


@app.get("/health")
async def health():
    return {"status": "ok"}


@app.post("/api/precheck")
async def precheck(
    manifest: Optional[UploadFile] = File(None),
    manifest_file: Optional[UploadFile] = File(None),
    file: Optional[UploadFile] = File(None),
    bill: Optional[UploadFile] = File(None),
    bill_file: Optional[UploadFile] = File(None),
    bl_file: Optional[UploadFile] = File(None),
    target_tax_amount: float = Form(...),
    target_item_count: int = Form(...),
    profile: str = Form("auto"),
    mode: str = Form("auto"),
):
    manifest_upload = pick_upload(manifest, manifest_file, file)
    bill_upload = pick_upload(bill, bill_file, bl_file)

    if not manifest_upload:
        raise HTTPException(status_code=400, detail="缺少清单 Excel")
    if not bill_upload:
        raise HTTPException(status_code=400, detail="缺少运输提单 PDF")

    manifest_name = manifest_upload.filename or ""
    bill_name = bill_upload.filename or ""
    if not manifest_name.lower().endswith((".xlsx", ".xls")):
        raise HTTPException(status_code=400, detail="清单仅支持 .xlsx / .xls")
    if not bill_name.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="运输提单仅支持 .pdf")

    manifest_path = save_upload(manifest_upload, "清单")
    bill_path = save_upload(bill_upload, "提单")
    precheck_id = uuid.uuid4().hex[:12]

    try:
        result = await build_precheck(
            manifest_path,
            bill_path,
            requested_profile="auto",
            target_tax_amount=target_tax_amount,
            target_item_count=target_item_count,
        )
    except Exception as exc:
        logger.exception("precheck failed")
        raise HTTPException(status_code=500, detail=f"预处理失败: {exc}") from exc

    record = {
        "precheck_id": precheck_id,
        "created_at": now_iso(),
        "manifest_path": str(manifest_path),
        "bill_path": str(bill_path),
        "precheck": result["precheck"],
        "manifest": result["manifest"],
        "bill": result["bill"],
        "query_cache": result["query_cache"],
        "input": result["input"],
    }
    write_json(precheck_path(precheck_id), record)
    return JSONResponse({"code": 200, "message": "预处理完成", "data": public_precheck_record(record)})


@app.post("/api/jobs")
async def create_job(precheck_id: str = Form(...)):
    record = read_json(precheck_path(precheck_id))
    precheck_data = record.get("precheck") or {}
    if not precheck_data.get("can_start"):
        reasons = "；".join(precheck_data.get("risk_reasons") or []) or "预处理判定不可执行"
        raise HTTPException(status_code=400, detail=f"预处理未通过，不能执行: {reasons}")

    task_id = uuid.uuid4().hex[:12]
    task_record = {
        "task_id": task_id,
        "precheck_id": precheck_id,
        "status": "queued",
        "phase": "queued",
        "progress": 0,
        "message": "等待后台任务启动",
        "created_at": now_iso(),
        "updated_at": now_iso(),
    }
    write_json(task_path(task_id), task_record)
    RUNNING_TASKS[task_id] = asyncio.create_task(run_clearance_job(task_id, record))
    return JSONResponse({"code": 200, "message": "任务已创建", "data": public_task_record(task_record)})


@app.get("/api/jobs/{task_id}")
async def get_job(task_id: str):
    return JSONResponse({"code": 200, "message": "ok", "data": public_task_record(read_json(task_path(task_id)))})


@app.post("/api/process")
async def process(
    manifest: Optional[UploadFile] = File(None),
    manifest_file: Optional[UploadFile] = File(None),
    file: Optional[UploadFile] = File(None),
    bill: Optional[UploadFile] = File(None),
    bill_file: Optional[UploadFile] = File(None),
    bl_file: Optional[UploadFile] = File(None),
    target_tax_amount: float = Form(...),
    target_item_count: int = Form(...),
    profile: str = Form("auto"),
    mode: str = Form("auto"),
):
    manifest_upload = pick_upload(manifest, manifest_file, file)
    bill_upload = pick_upload(bill, bill_file, bl_file)

    if not manifest_upload:
        raise HTTPException(status_code=400, detail="缺少清单 Excel")
    if not bill_upload:
        raise HTTPException(status_code=400, detail="缺少运输提单 PDF")

    manifest_name = manifest_upload.filename or ""
    bill_name = bill_upload.filename or ""
    if not manifest_name.lower().endswith((".xlsx", ".xls")):
        raise HTTPException(status_code=400, detail="清单仅支持 .xlsx / .xls")
    if not bill_name.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="运输提单仅支持 .pdf")

    manifest_path = save_upload(manifest_upload, "清单")
    bill_path = save_upload(bill_upload, "提单")
    requested_profile = "auto"

    try:
        result = await build_clearance(
            manifest_path,
            bill_path,
            OUTPUT_DIR,
            requested_profile,
            target_tax_amount=target_tax_amount,
            target_item_count=target_item_count,
        )
    except Exception as exc:
        logger.exception("process failed")
        raise HTTPException(status_code=500, detail=f"处理失败: {exc}") from exc

    output_name = Path(result["output_file"]).name
    logger.info(
        "process completed output=%s stats=%s flow=%s",
        output_name,
        result["stats"],
        result["flow"],
    )
    return JSONResponse(
        {
            "code": 200,
            "message": "处理成功",
            "data": {
                "task_id": uuid.uuid4().hex[:12],
                "stats": result["stats"],
                "flow": result["flow"],
                "manifest": result["manifest"],
                "bill": result["bill"],
                "bill_categories": result["bill_categories"],
                "output_rows": result["output_rows"],
                "filter_summary": result.get("filter_summary", {}),
                "download_url": f"/api/download/{output_name}",
            },
        }
    )


@app.get("/api/download/{filename}")
async def download(filename: str):
    path = OUTPUT_DIR / Path(filename).name
    if not path.exists():
        raise HTTPException(status_code=404, detail="文件不存在")
    return FileResponse(
        path=path,
        filename=path.name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


app.mount("/", StaticFiles(directory=STATIC_DIR, html=True), name="static")


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="0.0.0.0", port=8002)
