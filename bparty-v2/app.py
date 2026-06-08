from __future__ import annotations

import asyncio
import hashlib
import json
import logging
import os
import re
import secrets
import shutil
import time
import uuid
from datetime import datetime
from pathlib import Path
from typing import Optional

from fastapi import Body, Depends, FastAPI, File, Form, Header, HTTPException, Query, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from engine import build_clearance


BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
STATIC_DIR = BASE_DIR / "static"
RUNTIME_DIR = BASE_DIR / "runtime"
TASK_DIR = RUNTIME_DIR / "tasks"
CANCEL_DIR = RUNTIME_DIR / "cancelled_jobs"
ACCOUNT_PATH = RUNTIME_DIR / "accounts.json"
PERMISSION_PATH = RUNTIME_DIR / "permissions.json"
logger = logging.getLogger("bparty-v2")
UNFINISHED_STATUSES = {"queued", "running"}
FINISHED_STATUSES = {"succeeded", "failed"}
CLEANUP_INTERVAL_SECONDS = 24 * 60 * 60
PASSWORD_HASH_ITERATIONS = 260_000
TASK_SOURCE_FILES = {
    "manifest": {
        "key": "manifest_path",
        "label": "清单",
        "media_type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    },
    "bill": {
        "key": "bill_path",
        "label": "提单",
        "media_type": "application/pdf",
    },
}
DEFAULT_ACCOUNT_SEEDS = {
    "admin": {
        "password": os.getenv("BPARTY_ADMIN_PASSWORD", "admin123"),
        "role": "admin",
        "display_name": "管理员",
    },
    "user1": {
        "password": os.getenv("BPARTY_USER1_PASSWORD", "user123"),
        "role": "sub",
        "display_name": "子账号 1",
    },
    "user2": {
        "password": os.getenv("BPARTY_USER2_PASSWORD", "user123"),
        "role": "sub",
        "display_name": "子账号 2",
    },
    "user3": {
        "password": os.getenv("BPARTY_USER3_PASSWORD", "user123"),
        "role": "sub",
        "display_name": "子账号 3",
    },
    "user4": {
        "password": os.getenv("BPARTY_USER4_PASSWORD", "user123"),
        "role": "sub",
        "display_name": "子账号 4",
    },
}
DEFAULT_USER_PASSWORD = "user123"
AUTH_TOKENS: dict[str, str] = {}


def read_max_concurrent_jobs() -> int:
    try:
        return max(1, int(os.getenv("MAX_CONCURRENT_JOBS", "5")))
    except ValueError:
        logger.warning("invalid MAX_CONCURRENT_JOBS value; falling back to 5")
        return 5


def read_job_retention_days() -> int:
    try:
        return max(1, int(os.getenv("JOB_RETENTION_DAYS", "30")))
    except ValueError:
        logger.warning("invalid JOB_RETENTION_DAYS value; falling back to 30")
        return 30


UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
TASK_DIR.mkdir(parents=True, exist_ok=True)
CANCEL_DIR.mkdir(parents=True, exist_ok=True)
MAX_CONCURRENT_JOBS = read_max_concurrent_jobs()
JOB_RETENTION_DAYS = read_job_retention_days()
JOB_SEMAPHORE = asyncio.Semaphore(MAX_CONCURRENT_JOBS)
RUNNING_TASKS: dict[str, asyncio.Task] = {}
CANCELLED_TASK_IDS: set[str] = set()
RETENTION_CLEANUP_TASK: asyncio.Task | None = None

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


@app.on_event("startup")
async def start_retention_cleanup() -> None:
    global RETENTION_CLEANUP_TASK
    cleanup_expired_job_records()
    RETENTION_CLEANUP_TASK = asyncio.create_task(retention_cleanup_loop())


@app.on_event("shutdown")
async def stop_retention_cleanup() -> None:
    global RETENTION_CLEANUP_TASK
    if RETENTION_CLEANUP_TASK:
        RETENTION_CLEANUP_TASK.cancel()
        await asyncio.gather(RETENTION_CLEANUP_TASK, return_exceptions=True)
        RETENTION_CLEANUP_TASK = None


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


def now_epoch() -> float:
    return time.time()


def parse_record_epoch(
    record: dict,
    keys: tuple[str, ...] = ("completed_at_epoch", "completed_at", "updated_at", "created_at"),
) -> float | None:
    for key in keys:
        value = record.get(key)
        if isinstance(value, (int, float)):
            return float(value)
        if not value:
            continue
        text = str(value).strip()
        if not text:
            continue
        try:
            normalized = text.replace("Z", "+00:00")
            return datetime.fromisoformat(normalized).timestamp()
        except ValueError:
            continue
    return None


def write_json(path: Path, payload: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_suffix(path.suffix + ".tmp")
    temp.write_text(json.dumps(payload, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    temp.replace(path)


def read_json(path: Path) -> dict:
    if not path.exists():
        raise HTTPException(status_code=404, detail="记录不存在或已过期")
    return json.loads(path.read_text(encoding="utf-8"))


def hash_password(password: str, salt: str | None = None) -> str:
    salt_value = salt or secrets.token_hex(16)
    digest = hashlib.pbkdf2_hmac(
        "sha256",
        password.encode("utf-8"),
        salt_value.encode("utf-8"),
        PASSWORD_HASH_ITERATIONS,
    ).hex()
    return f"pbkdf2_sha256${PASSWORD_HASH_ITERATIONS}${salt_value}${digest}"


def verify_password(password: str, stored_password: str) -> bool:
    parts = str(stored_password or "").split("$")
    if len(parts) != 4 or parts[0] != "pbkdf2_sha256":
        return False
    try:
        iterations = int(parts[1])
    except ValueError:
        return False
    salt = parts[2]
    expected = parts[3]
    digest = hashlib.pbkdf2_hmac(
        "sha256",
        password.encode("utf-8"),
        salt.encode("utf-8"),
        iterations,
    ).hex()
    return secrets.compare_digest(digest, expected)


def seed_account_record(user_id: str, payload: dict) -> dict:
    return {
        "user_id": user_id,
        "role": payload["role"],
        "display_name": payload["display_name"],
        "password_hash": hash_password(str(payload["password"])),
        "created_at": now_iso(),
        "updated_at": now_iso(),
    }


def normalize_account_record(user_id: str, payload: dict) -> dict | None:
    if user_id != "admin" and not re.fullmatch(r"user[1-9][0-9]*", user_id):
        return None
    role = "admin" if user_id == "admin" else "sub"
    display_name = str(payload.get("display_name") or ("管理员" if role == "admin" else f"子账号 {user_id[4:]}"))
    password_hash = str(payload.get("password_hash") or "")
    if not password_hash:
        password = str(payload.get("password") or DEFAULT_ACCOUNT_SEEDS.get(user_id, {}).get("password") or DEFAULT_USER_PASSWORD)
        password_hash = hash_password(password)
    return {
        "user_id": user_id,
        "role": role,
        "display_name": display_name,
        "password_hash": password_hash,
        "created_at": str(payload.get("created_at") or now_iso()),
        "updated_at": str(payload.get("updated_at") or now_iso()),
    }


def read_accounts() -> dict:
    changed = False
    accounts: dict[str, dict] = {}
    if ACCOUNT_PATH.exists():
        try:
            raw = json.loads(ACCOUNT_PATH.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            logger.warning("accounts file is unreadable; resetting defaults path=%s", ACCOUNT_PATH)
            raw = {}
            changed = True
        raw_accounts = raw.get("accounts") if isinstance(raw, dict) else None
        if not isinstance(raw_accounts, dict):
            raw_accounts = {}
            changed = True
        for user_id, payload in raw_accounts.items():
            if not isinstance(payload, dict):
                changed = True
                continue
            record = normalize_account_record(str(user_id), payload)
            if record is None:
                changed = True
                continue
            accounts[record["user_id"]] = record

    for user_id, payload in DEFAULT_ACCOUNT_SEEDS.items():
        if user_id not in accounts:
            accounts[user_id] = seed_account_record(user_id, payload)
            changed = True

    if changed or not ACCOUNT_PATH.exists():
        write_accounts(accounts)
    return accounts


def write_accounts(accounts: dict) -> None:
    cleaned: dict[str, dict] = {}
    for user_id, payload in accounts.items():
        if not isinstance(payload, dict):
            continue
        if str(payload.get("password_hash") or ""):
            record = normalize_account_record(str(user_id), payload)
            if record is not None:
                cleaned[record["user_id"]] = record
    write_json(ACCOUNT_PATH, {"accounts": cleaned})


def sub_account_ids(accounts: dict | None = None) -> tuple[str, ...]:
    source = accounts or read_accounts()
    return tuple(
        sorted(
            (user_id for user_id, account in source.items() if account.get("role") == "sub"),
            key=lambda user_id: int(user_id[4:]) if re.fullmatch(r"user[1-9][0-9]*", user_id) else 0,
        )
    )


def public_account_record(user_id: str, account: dict, permissions: dict | None = None) -> dict:
    permission_source = permissions if permissions is not None else read_permissions()
    can_view_all = account.get("role") == "admin" or bool(
        permission_source.get(user_id, {}).get("can_view_all_tasks")
    )
    return {
        "user_id": user_id,
        "role": account["role"],
        "display_name": account["display_name"],
        "can_view_all_tasks": can_view_all,
    }


def invalidate_user_tokens(user_id: str) -> None:
    for token, token_user_id in list(AUTH_TOKENS.items()):
        if token_user_id == user_id:
            AUTH_TOKENS.pop(token, None)


def next_user_id(accounts: dict) -> str:
    max_index = 0
    for user_id, account in accounts.items():
        if account.get("role") != "sub":
            continue
        match = re.fullmatch(r"user([1-9][0-9]*)", user_id)
        if match:
            max_index = max(max_index, int(match.group(1)))
    return f"user{max_index + 1}"


def set_account_password(user_id: str, new_password: str) -> dict:
    password = str(new_password or "")
    if not password:
        raise HTTPException(status_code=400, detail="新密码不能为空")
    accounts = read_accounts()
    account = accounts.get(user_id)
    if not account:
        raise HTTPException(status_code=404, detail="账号不存在")
    account["password_hash"] = hash_password(password)
    account["updated_at"] = now_iso()
    accounts[user_id] = account
    write_accounts(accounts)
    invalidate_user_tokens(user_id)
    return public_account_record(user_id, account)


def default_permissions() -> dict:
    return {user_id: {"can_view_all_tasks": False} for user_id in sub_account_ids()}


def read_permissions() -> dict:
    permissions = default_permissions()
    if not PERMISSION_PATH.exists():
        write_json(PERMISSION_PATH, permissions)
        return permissions
    try:
        raw = json.loads(PERMISSION_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        logger.warning("permissions file is unreadable; resetting defaults path=%s", PERMISSION_PATH)
        write_json(PERMISSION_PATH, permissions)
        return permissions
    if not isinstance(raw, dict):
        write_json(PERMISSION_PATH, permissions)
        return permissions
    for user_id in sub_account_ids():
        value = raw.get(user_id)
        if isinstance(value, dict):
            permissions[user_id]["can_view_all_tasks"] = bool(value.get("can_view_all_tasks"))
    return permissions


def write_permissions(permissions: dict) -> None:
    cleaned = default_permissions()
    for user_id in sub_account_ids():
        value = permissions.get(user_id)
        if isinstance(value, dict):
            cleaned[user_id]["can_view_all_tasks"] = bool(value.get("can_view_all_tasks"))
    write_json(PERMISSION_PATH, cleaned)


def public_user(user_id: str) -> dict:
    accounts = read_accounts()
    return public_account_record(user_id, accounts[user_id], read_permissions())


def resolve_token(authorization: Optional[str], x_auth_token: Optional[str], token: Optional[str]) -> str:
    if authorization:
        scheme, _, value = authorization.partition(" ")
        if scheme.lower() == "bearer" and value.strip():
            return value.strip()
    if x_auth_token:
        return x_auth_token.strip()
    return (token or "").strip()


def get_current_user(
    authorization: Optional[str] = Header(None),
    x_auth_token: Optional[str] = Header(None, alias="X-Auth-Token"),
    token: Optional[str] = Query(None),
) -> dict:
    auth_token = resolve_token(authorization, x_auth_token, token)
    user_id = AUTH_TOKENS.get(auth_token)
    if not user_id or user_id not in read_accounts():
        raise HTTPException(status_code=401, detail="请先登录")
    return public_user(user_id)


def require_admin(current_user: dict = Depends(get_current_user)) -> dict:
    if current_user.get("role") != "admin":
        raise HTTPException(status_code=403, detail="只有 admin 可以管理权限")
    return current_user


def user_can_view_all_tasks(current_user: dict) -> bool:
    return current_user.get("role") == "admin" or bool(current_user.get("can_view_all_tasks"))


def user_can_access_task(record: dict, current_user: dict) -> bool:
    if user_can_view_all_tasks(current_user):
        return True
    return str(record.get("owner_user_id") or "") == str(current_user.get("user_id") or "")


def ensure_task_access(record: dict, current_user: dict) -> None:
    if not user_can_access_task(record, current_user):
        raise HTTPException(status_code=404, detail="记录不存在或无权访问")


def user_can_manage_task(record: dict, current_user: dict) -> bool:
    if current_user.get("role") == "admin":
        return True
    return str(record.get("owner_user_id") or "") == str(current_user.get("user_id") or "")


def ensure_task_management_access(record: dict, current_user: dict) -> None:
    if not user_can_manage_task(record, current_user):
        raise HTTPException(status_code=403, detail="只能取消或删除自己的任务")


def find_accessible_task_by_output(filename: str, current_user: dict) -> dict | None:
    safe_filename = Path(filename).name
    for record in list_task_records(None):
        output_file = record.get("output_file")
        if not output_file:
            continue
        if Path(str(output_file)).name != safe_filename:
            continue
        if user_can_access_task(record, current_user):
            return record
    return None


def resolve_task_source_file(record: dict, kind: str) -> Path | None:
    source_config = TASK_SOURCE_FILES.get(kind)
    if source_config is None:
        return None
    input_data = record.get("input") if isinstance(record.get("input"), dict) else {}
    value = record.get(source_config["key"]) or input_data.get(source_config["key"])
    if not value:
        return None
    path = Path(str(value))
    if not path.is_absolute():
        path = UPLOAD_DIR / path.name
    try:
        upload_dir = UPLOAD_DIR.resolve()
        resolved_path = path.resolve()
    except OSError:
        return None
    if resolved_path.parent != upload_dir:
        return None
    return resolved_path


def task_source_files_payload(record: dict) -> dict:
    if str(record.get("status") or "") != "failed":
        return {}
    task_id = str(record.get("task_id") or "")
    if not task_id:
        return {}
    source_files: dict[str, dict[str, str]] = {}
    for kind, source_config in TASK_SOURCE_FILES.items():
        path = resolve_task_source_file(record, kind)
        if path is None or not path.exists():
            continue
        source_files[kind] = {
            "label": str(source_config["label"]),
            "filename": path.name,
            "download_url": f"/api/task-files/{Path(task_id).name}/{kind}",
        }
    return source_files


def task_path(task_id: str) -> Path:
    return TASK_DIR / f"{Path(task_id).name}.json"


def cancel_marker_path(task_id: str) -> Path:
    return CANCEL_DIR / f"{Path(task_id).name}.cancelled"


def is_task_cancelled(task_id: str) -> bool:
    safe_task_id = Path(task_id).name
    return safe_task_id in CANCELLED_TASK_IDS or cancel_marker_path(safe_task_id).exists()


def mark_task_cancelled(task_id: str) -> None:
    safe_task_id = Path(task_id).name
    CANCELLED_TASK_IDS.add(safe_task_id)
    cancel_marker_path(safe_task_id).parent.mkdir(parents=True, exist_ok=True)
    cancel_marker_path(safe_task_id).write_text(now_iso(), encoding="utf-8")


def clear_task_cancelled(task_id: str) -> None:
    safe_task_id = Path(task_id).name
    CANCELLED_TASK_IDS.discard(safe_task_id)
    try:
        cancel_marker_path(safe_task_id).unlink()
    except FileNotFoundError:
        pass


def safe_unlink(path: Path) -> bool:
    try:
        path.unlink()
        return True
    except FileNotFoundError:
        return False


def cleanup_task_artifacts(record: dict, *, delete_inputs: bool = True, delete_output: bool = True) -> dict:
    removed: list[str] = []
    if delete_output and record.get("output_file"):
        output_path = Path(str(record["output_file"]))
        if not output_path.is_absolute():
            output_path = OUTPUT_DIR / output_path.name
        if output_path.parent.resolve() == OUTPUT_DIR.resolve() and safe_unlink(output_path):
            removed.append(str(output_path))

    if delete_inputs:
        candidate_paths: list[str] = []
        input_data = record.get("input") if isinstance(record.get("input"), dict) else {}
        for key in ("manifest_path", "bill_path"):
            value = record.get(key) or input_data.get(key)
            if value:
                candidate_paths.append(str(value))
        for value in candidate_paths:
            upload_path = Path(value)
            if not upload_path.is_absolute():
                upload_path = UPLOAD_DIR / upload_path.name
            if upload_path.parent.resolve() == UPLOAD_DIR.resolve() and safe_unlink(upload_path):
                removed.append(str(upload_path))

    return {"removed_files": removed}


def task_sort_key(record: dict) -> tuple[int, str, str]:
    status_priority = {"running": 0, "queued": 1}.get(str(record.get("status") or ""), 2)
    return (status_priority, str(record.get("created_at") or ""), str(record.get("task_id") or ""))


def public_task_record(record: dict) -> dict:
    payload = dict(record)
    payload["display_task_no"] = str(payload.get("display_task_no") or payload.get("task_no") or payload.get("task_id") or "")
    started_at_epoch = payload.get("started_at_epoch")
    if isinstance(started_at_epoch, (int, float)):
        completed_at_epoch = payload.get("completed_at_epoch")
        end_epoch = completed_at_epoch if isinstance(completed_at_epoch, (int, float)) else now_epoch()
        elapsed_seconds = max(0.0, end_epoch - started_at_epoch)
        payload["elapsed_seconds"] = round(elapsed_seconds, 2)
        payload["elapsed_minutes"] = round(elapsed_seconds / 60, 2)
    if payload.get("output_file"):
        payload["download_url"] = f"/api/download/{Path(payload['output_file']).name}"
    source_files = task_source_files_payload(record)
    if source_files:
        payload["source_files"] = source_files
    payload.pop("manifest_path", None)
    payload.pop("bill_path", None)
    payload.pop("input", None)
    payload.pop("query_cache", None)
    payload.pop("started_at_epoch", None)
    payload.pop("completed_at_epoch", None)
    return payload


def update_task(task_id: str, **updates) -> dict:
    if is_task_cancelled(task_id):
        return {"task_id": task_id, "status": "cancelled"}
    path = task_path(task_id)
    try:
        record = read_json(path)
    except HTTPException:
        record = {"task_id": task_id}
    record.update(updates)
    started_at_epoch = record.get("started_at_epoch")
    if isinstance(started_at_epoch, (int, float)):
        completed_at_epoch = record.get("completed_at_epoch")
        end_epoch = completed_at_epoch if isinstance(completed_at_epoch, (int, float)) else now_epoch()
        elapsed_seconds = max(0.0, end_epoch - started_at_epoch)
        record["elapsed_seconds"] = round(elapsed_seconds, 2)
        record["elapsed_minutes"] = round(elapsed_seconds / 60, 2)
    record["updated_at"] = now_iso()
    write_json(path, record)
    return record


def add_task_feedback(task_id: str, payload: dict, current_user: dict) -> dict:
    safe_task_id = Path(task_id).name
    record = read_json(task_path(safe_task_id))
    ensure_task_access(record, current_user)
    content = str(payload.get("content") or "").strip()
    if not content:
        raise HTTPException(status_code=400, detail="反馈内容不能为空")
    if len(content) > 2000:
        raise HTTPException(status_code=400, detail="反馈内容不能超过 2000 字")
    feedback_type = str(payload.get("feedback_type") or "other").strip() or "other"
    allowed_types = {"price", "hs", "quantity_weight", "category", "result", "other"}
    if feedback_type not in allowed_types:
        feedback_type = "other"
    scope = str(payload.get("scope") or "task").strip() or "task"
    if scope not in {"task", "row"}:
        scope = "task"
    row_index = payload.get("row_index")
    if row_index in ("", None):
        row_index = None
    else:
        try:
            row_index = max(1, int(row_index))
        except (TypeError, ValueError):
            row_index = None
    entry = {
        "feedback_id": uuid.uuid4().hex[:12],
        "task_id": safe_task_id,
        "created_at": now_iso(),
        "user_id": current_user["user_id"],
        "user_display_name": current_user.get("display_name") or current_user["user_id"],
        "feedback_type": feedback_type,
        "scope": scope,
        "row_index": row_index,
        "row_label": str(payload.get("row_label") or "").strip()[:200],
        "content": content,
    }
    feedbacks = record.get("feedbacks")
    if not isinstance(feedbacks, list):
        feedbacks = []
    feedbacks.append(entry)
    record["feedbacks"] = feedbacks
    record["feedback_count"] = len(feedbacks)
    record["updated_at"] = now_iso()
    write_json(task_path(safe_task_id), record)
    return record


def list_task_records(statuses: set[str] | None = None) -> list[dict]:
    records: list[dict] = []
    for path in TASK_DIR.glob("*.json"):
        try:
            record = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            logger.warning("skipping unreadable task record path=%s", path)
            continue
        if is_task_cancelled(str(record.get("task_id") or path.stem)):
            continue
        status = str(record.get("status") or "")
        if statuses is not None and status not in statuses:
            continue
        records.append(record)
    return sorted(records, key=task_sort_key)


def build_job_queue_payload(
    include_finished: bool = False,
    limit: int = 20,
    current_user: dict | None = None,
) -> dict:
    cleanup_expired_job_records()
    records = list_task_records(None if include_finished else {"queued", "running"})
    if current_user is not None and not user_can_view_all_tasks(current_user):
        records = [
            record
            for record in records
            if str(record.get("owner_user_id") or "") == str(current_user.get("user_id") or "")
        ]
    running_count = sum(1 for record in records if record.get("status") == "running")
    queued_count = sum(1 for record in records if record.get("status") == "queued")
    unfinished: list[dict] = []
    finished: list[dict] = []
    jobs: list[dict] = []
    running_position = 0
    queue_position = 0
    for record in records:
        job = public_task_record(record)
        if record.get("status") == "running":
            running_position += 1
            job["running_position"] = running_position
        elif record.get("status") == "queued":
            queue_position += 1
            job["queue_position"] = queue_position
        if record.get("status") in {"queued", "running"}:
            unfinished.append(job)
        else:
            finished.append(job)
    jobs.extend(unfinished)
    if include_finished:
        finished = sorted(
            finished,
            key=lambda job: str(job.get("completed_at") or job.get("updated_at") or ""),
            reverse=True,
        )
        jobs.extend(finished[: max(0, limit)])
    return {
        "max_concurrency": MAX_CONCURRENT_JOBS,
        "running_count": running_count,
        "queued_count": queued_count,
        "jobs": jobs,
    }


def delete_task_record(task_id: str) -> dict:
    safe_task_id = Path(task_id).name
    path = task_path(safe_task_id)
    record = read_json(path)
    status = str(record.get("status") or "")
    if status in UNFINISHED_STATUSES:
        raise HTTPException(status_code=409, detail="任务仍在排队或运行，请使用取消任务")
    if status not in FINISHED_STATUSES:
        raise HTTPException(status_code=409, detail="任务状态不支持删除")
    artifacts = cleanup_task_artifacts(record)
    safe_unlink(path)
    clear_task_cancelled(safe_task_id)
    return {
        "task_id": safe_task_id,
        "status": "deleted",
        "removed_files": artifacts["removed_files"],
    }


def cleanup_expired_job_records(retention_days: int | None = None, current_epoch: float | None = None) -> dict:
    retention = retention_days or JOB_RETENTION_DAYS
    cutoff_epoch = (current_epoch if current_epoch is not None else now_epoch()) - retention * 24 * 60 * 60
    deleted_records = 0
    removed_files: list[str] = []
    for path in TASK_DIR.glob("*.json"):
        try:
            record = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            logger.warning("skipping unreadable task record path=%s", path)
            continue
        status = str(record.get("status") or "")
        if status in UNFINISHED_STATUSES:
            continue
        record_epoch = parse_record_epoch(record)
        if record_epoch is None or record_epoch >= cutoff_epoch:
            continue
        artifacts = cleanup_task_artifacts(record)
        removed_files.extend(artifacts["removed_files"])
        if safe_unlink(path):
            deleted_records += 1
        clear_task_cancelled(str(record.get("task_id") or path.stem))

    deleted_markers = cleanup_expired_cancel_markers(retention, current_epoch=current_epoch)
    return {
        "retention_days": retention,
        "deleted_records": deleted_records,
        "deleted_cancel_markers": deleted_markers,
        "removed_files": removed_files,
    }


def cleanup_expired_cancel_markers(retention_days: int, current_epoch: float | None = None) -> int:
    cutoff_epoch = (current_epoch if current_epoch is not None else now_epoch()) - retention_days * 24 * 60 * 60
    deleted = 0
    for path in CANCEL_DIR.glob("*.cancelled"):
        try:
            marker_epoch = path.stat().st_mtime
        except OSError:
            continue
        if marker_epoch >= cutoff_epoch:
            continue
        CANCELLED_TASK_IDS.discard(path.stem)
        if safe_unlink(path):
            deleted += 1
    return deleted


async def retention_cleanup_loop() -> None:
    while True:
        try:
            cleanup_expired_job_records()
        except Exception:
            logger.exception("retention cleanup failed")
        await asyncio.sleep(CLEANUP_INTERVAL_SECONDS)


def schedule_clearance_job(task_id: str, task_record: dict) -> asyncio.Task:
    clear_task_cancelled(task_id)
    task = asyncio.create_task(run_clearance_job(task_id, task_record))
    RUNNING_TASKS[task_id] = task

    def cleanup_task(finished_task: asyncio.Task) -> None:
        RUNNING_TASKS.pop(task_id, None)
        if finished_task.cancelled():
            return
        try:
            exc = finished_task.exception()
        except asyncio.CancelledError:
            return
        if exc:
            logger.error(
                "background job task crashed task_id=%s",
                task_id,
                exc_info=(type(exc), exc, exc.__traceback__),
            )

    task.add_done_callback(cleanup_task)
    return task


def build_processing_input(
    manifest: Optional[UploadFile],
    manifest_file: Optional[UploadFile],
    file: Optional[UploadFile],
    bill: Optional[UploadFile],
    bill_file: Optional[UploadFile],
    bl_file: Optional[UploadFile],
    target_tax_amount: float,
    target_item_count: int,
) -> dict:
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
    return {
        "manifest_path": str(manifest_path),
        "bill_path": str(bill_path),
        "requested_profile": "auto",
        "target_tax_amount": target_tax_amount,
        "target_item_count": target_item_count,
    }


async def run_clearance_job(task_id: str, task_record: dict) -> None:
    input_data = task_record["input"]
    query_cache = task_record.get("query_cache") or {"product": {}, "hs": {}, "bill": {}}
    try:
        async with JOB_SEMAPHORE:
            if is_task_cancelled(task_id):
                return
            started_at_epoch = now_epoch()
            update_task(
                task_id,
                status="running",
                phase="queued",
                progress=1,
                message="任务已启动",
                started_at=now_iso(),
                started_at_epoch=started_at_epoch,
                manifest_path=input_data["manifest_path"],
                bill_path=input_data["bill_path"],
            )

            async def progress_callback(payload: dict) -> None:
                if is_task_cancelled(task_id):
                    return
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
                if is_task_cancelled(task_id):
                    return
                completed_at_epoch = now_epoch()
                elapsed_seconds = round(max(0.0, completed_at_epoch - started_at_epoch), 2)
                stats = dict(result["stats"])
                stats["task_elapsed_seconds"] = elapsed_seconds
                stats["task_elapsed_minutes"] = round(elapsed_seconds / 60, 2)
                update_task(
                    task_id,
                    status="succeeded",
                    phase="done",
                    progress=100,
                    message="处理完成",
                    completed_at=now_iso(),
                    completed_at_epoch=completed_at_epoch,
                    stats=stats,
                    flow=result["flow"],
                    manifest=result["manifest"],
                    bill=result["bill"],
                    bill_categories=result["bill_categories"],
                    output_rows=result["output_rows"],
                    filter_summary=result.get("filter_summary", {}),
                    output_file=result["output_file"],
                )
            except asyncio.CancelledError:
                raise
            except Exception as exc:
                if is_task_cancelled(task_id):
                    return
                logger.exception("job failed task_id=%s", task_id)
                completed_at_epoch = now_epoch()
                update_task(
                    task_id,
                    status="failed",
                    phase="failed",
                    progress=100,
                    message=str(exc),
                    error=str(exc),
                    completed_at=now_iso(),
                    completed_at_epoch=completed_at_epoch,
                )
    finally:
        if is_task_cancelled(task_id):
            try:
                task_path(task_id).unlink()
            except FileNotFoundError:
                pass


async def cancel_job_record(task_id: str) -> dict:
    path = task_path(task_id)
    if not path.exists():
        raise HTTPException(status_code=404, detail="任务不存在或已结束")
    record = read_json(path)
    if record.get("status") not in {"queued", "running"}:
        raise HTTPException(status_code=409, detail="任务已结束，不能取消")
    mark_task_cancelled(task_id)
    task = RUNNING_TASKS.pop(task_id, None)
    if task and not task.done():
        task.cancel()
    try:
        path.unlink()
    except FileNotFoundError:
        pass
    if task:
        await asyncio.sleep(0)
    return {"task_id": task_id, "status": "cancelled"}


@app.get("/health")
async def health():
    return {"status": "ok"}


@app.post("/api/login")
async def login(payload: dict = Body(...)):
    username = str(payload.get("username") or "").strip()
    password = str(payload.get("password") or "")
    account = read_accounts().get(username)
    if not account or not verify_password(password, str(account.get("password_hash") or "")):
        raise HTTPException(status_code=401, detail="账号或密码不正确")
    token = secrets.token_urlsafe(32)
    AUTH_TOKENS[token] = username
    return JSONResponse(
        {
            "code": 200,
            "message": "登录成功",
            "data": {
                "token": token,
                "user": public_user(username),
            },
        }
    )


@app.post("/api/logout")
async def logout(
    authorization: Optional[str] = Header(None),
    x_auth_token: Optional[str] = Header(None, alias="X-Auth-Token"),
    token: Optional[str] = Query(None),
):
    auth_token = resolve_token(authorization, x_auth_token, token)
    if auth_token:
        AUTH_TOKENS.pop(auth_token, None)
    return JSONResponse({"code": 200, "message": "已退出", "data": {"status": "ok"}})


@app.get("/api/me")
async def me(current_user: dict = Depends(get_current_user)):
    return JSONResponse({"code": 200, "message": "ok", "data": current_user})


@app.get("/api/accounts")
async def get_accounts(_: dict = Depends(require_admin)):
    accounts = read_accounts()
    permissions = read_permissions()
    users = [
        public_account_record(user_id, accounts[user_id], permissions)
        for user_id in sorted(accounts, key=lambda item: (0 if item == "admin" else 1, int(item[4:]) if item.startswith("user") and item[4:].isdigit() else 0))
    ]
    return JSONResponse({"code": 200, "message": "ok", "data": {"users": users}})


@app.post("/api/accounts")
async def create_account(payload: dict = Body(default={}), _: dict = Depends(require_admin)):
    accounts = read_accounts()
    user_id = next_user_id(accounts)
    password = str(payload.get("password") or DEFAULT_USER_PASSWORD)
    account = {
        "user_id": user_id,
        "role": "sub",
        "display_name": f"子账号 {user_id[4:]}",
        "password_hash": hash_password(password),
        "created_at": now_iso(),
        "updated_at": now_iso(),
    }
    accounts[user_id] = account
    write_accounts(accounts)
    permissions = read_permissions()
    permissions.setdefault(user_id, {"can_view_all_tasks": False})
    write_permissions(permissions)
    return JSONResponse(
        {
            "code": 200,
            "message": "账号已创建",
            "data": public_account_record(user_id, account, read_permissions()),
        }
    )


@app.patch("/api/accounts/{user_id}/password")
async def update_account_password(
    user_id: str,
    payload: dict = Body(...),
    current_user: dict = Depends(require_admin),
):
    safe_user_id = Path(user_id).name
    accounts = read_accounts()
    account = accounts.get(safe_user_id)
    if not account:
        raise HTTPException(status_code=404, detail="账号不存在")
    if safe_user_id == "admin":
        current_password = str(payload.get("current_admin_password") or "")
        admin_account = accounts.get("admin") or {}
        if not current_password or not verify_password(current_password, str(admin_account.get("password_hash") or "")):
            raise HTTPException(status_code=403, detail="当前 admin 密码不正确")
    updated = set_account_password(safe_user_id, str(payload.get("new_password") or ""))
    return JSONResponse(
        {
            "code": 200,
            "message": "密码已更新",
            "data": {
                "user": updated,
                "current_user_logged_out": safe_user_id == current_user["user_id"],
            },
        }
    )


@app.get("/api/permissions")
async def get_permissions(_: dict = Depends(require_admin)):
    permissions = read_permissions()
    users = []
    accounts = read_accounts()
    for user_id in sub_account_ids(accounts):
        users.append(public_account_record(user_id, accounts[user_id], permissions))
    return JSONResponse({"code": 200, "message": "ok", "data": {"users": users}})


@app.patch("/api/permissions/{user_id}")
async def update_permission(
    user_id: str,
    payload: dict = Body(...),
    _: dict = Depends(require_admin),
):
    safe_user_id = Path(user_id).name
    accounts = read_accounts()
    if safe_user_id not in sub_account_ids(accounts):
        raise HTTPException(status_code=404, detail="子账号不存在")
    permissions = read_permissions()
    permissions[safe_user_id]["can_view_all_tasks"] = bool(payload.get("can_view_all_tasks"))
    write_permissions(permissions)
    return JSONResponse(
        {
            "code": 200,
            "message": "权限已更新",
            "data": {
                "user_id": safe_user_id,
                "can_view_all_tasks": permissions[safe_user_id]["can_view_all_tasks"],
            },
        }
    )


@app.post("/api/jobs")
async def create_job(
    manifest: Optional[UploadFile] = File(None),
    manifest_file: Optional[UploadFile] = File(None),
    file: Optional[UploadFile] = File(None),
    bill: Optional[UploadFile] = File(None),
    bill_file: Optional[UploadFile] = File(None),
    bl_file: Optional[UploadFile] = File(None),
    target_tax_amount: float = Form(...),
    target_item_count: int = Form(...),
    display_task_no: str = Form(...),
    task_no: str = Form(""),
    profile: str = Form("auto"),
    mode: str = Form("auto"),
    current_user: dict = Depends(get_current_user),
):
    display_no = (display_task_no or task_no or "").strip()
    if not display_no:
        raise HTTPException(status_code=400, detail="缺少任务编号")
    if len(display_no) > 120:
        raise HTTPException(status_code=400, detail="任务编号不能超过 120 个字符")
    input_data = build_processing_input(
        manifest,
        manifest_file,
        file,
        bill,
        bill_file,
        bl_file,
        target_tax_amount,
        target_item_count,
    )
    task_id = uuid.uuid4().hex[:12]
    task_record = {
        "task_id": task_id,
        "display_task_no": display_no,
        "owner_user_id": current_user["user_id"],
        "owner_display_name": current_user["display_name"],
        "status": "queued",
        "phase": "queued",
        "progress": 0,
        "message": "等待后台任务启动",
        "created_at": now_iso(),
        "updated_at": now_iso(),
        "input": input_data,
        "query_cache": {"product": {}, "hs": {}, "bill": {}},
    }
    write_json(task_path(task_id), task_record)
    schedule_clearance_job(task_id, task_record)
    return JSONResponse({"code": 200, "message": "任务已创建", "data": public_task_record(task_record)})


@app.get("/api/jobs")
async def list_jobs(
    include_finished: bool = False,
    limit: int = 20,
    current_user: dict = Depends(get_current_user),
):
    return JSONResponse(
        {"code": 200, "message": "ok", "data": build_job_queue_payload(include_finished, limit, current_user)}
    )


@app.get("/api/jobs/{task_id}")
async def get_job(task_id: str, current_user: dict = Depends(get_current_user)):
    record = read_json(task_path(task_id))
    ensure_task_access(record, current_user)
    return JSONResponse({"code": 200, "message": "ok", "data": public_task_record(record)})


@app.post("/api/jobs/{task_id}/feedback")
async def submit_job_feedback(
    task_id: str,
    payload: dict = Body(...),
    current_user: dict = Depends(get_current_user),
):
    record = add_task_feedback(task_id, payload, current_user)
    return JSONResponse({"code": 200, "message": "反馈已记录", "data": public_task_record(record)})


@app.get("/api/task-files/{task_id}/{kind}")
async def download_task_source_file(task_id: str, kind: str, current_user: dict = Depends(get_current_user)):
    if kind not in TASK_SOURCE_FILES:
        raise HTTPException(status_code=404, detail="文件不存在")
    record = read_json(task_path(task_id))
    ensure_task_access(record, current_user)
    if str(record.get("status") or "") != "failed":
        raise HTTPException(status_code=404, detail="文件不存在")
    path = resolve_task_source_file(record, kind)
    if path is None or not path.exists():
        raise HTTPException(status_code=404, detail="文件不存在")
    source_config = TASK_SOURCE_FILES[kind]
    return FileResponse(path=path, filename=path.name, media_type=str(source_config["media_type"]))


@app.delete("/api/jobs/{task_id}")
async def cancel_job(task_id: str, current_user: dict = Depends(get_current_user)):
    record = read_json(task_path(task_id))
    ensure_task_access(record, current_user)
    ensure_task_management_access(record, current_user)
    status = str(record.get("status") or "")
    if status in UNFINISHED_STATUSES:
        return JSONResponse({"code": 200, "message": "任务已取消", "data": await cancel_job_record(task_id)})
    return JSONResponse({"code": 200, "message": "任务记录已删除", "data": delete_task_record(task_id)})


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
    current_user: dict = Depends(get_current_user),
):
    input_data = build_processing_input(
        manifest,
        manifest_file,
        file,
        bill,
        bill_file,
        bl_file,
        target_tax_amount,
        target_item_count,
    )

    try:
        async with JOB_SEMAPHORE:
            result = await build_clearance(
                input_data["manifest_path"],
                input_data["bill_path"],
                OUTPUT_DIR,
                input_data["requested_profile"],
                target_tax_amount=target_tax_amount,
                target_item_count=target_item_count,
            )
    except Exception as exc:
        logger.exception("process failed")
        raise HTTPException(status_code=500, detail=f"处理失败: {exc}") from exc

    output_name = Path(result["output_file"]).name
    task_id = uuid.uuid4().hex[:12]
    completed_at = now_iso()
    task_record = {
        "task_id": task_id,
        "display_task_no": task_id,
        "owner_user_id": current_user["user_id"],
        "status": "succeeded",
        "phase": "done",
        "progress": 100,
        "message": "处理成功",
        "created_at": completed_at,
        "updated_at": completed_at,
        "completed_at": completed_at,
        "completed_at_epoch": now_epoch(),
        "input": input_data,
        "query_cache": {"product": {}, "hs": {}, "bill": {}},
        **result,
    }
    write_json(task_path(task_id), task_record)
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
            "data": public_task_record(task_record),
        }
    )


@app.get("/api/download/{filename}")
async def download(filename: str, current_user: dict = Depends(get_current_user)):
    record = find_accessible_task_by_output(filename, current_user)
    if record is None:
        raise HTTPException(status_code=404, detail="文件不存在或无权下载")
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
