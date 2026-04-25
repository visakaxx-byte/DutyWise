"""
工具函数
"""
import uuid
from datetime import datetime
from pathlib import Path
from typing import Optional
import logging

logger = logging.getLogger(__name__)


def generate_shipment_no() -> str:
    """生成票据编号"""
    timestamp = datetime.now().strftime('%Y%m%d')
    random_suffix = str(uuid.uuid4())[:6].upper()
    return f"SHIP{timestamp}{random_suffix}"


def generate_task_id() -> str:
    """生成任务ID"""
    return f"task_{uuid.uuid4().hex[:12]}"


def ensure_dir(directory: str) -> Path:
    """确保目录存在"""
    path = Path(directory)
    path.mkdir(parents=True, exist_ok=True)
    return path


def get_file_extension(filename: str) -> str:
    """获取文件扩展名"""
    return Path(filename).suffix.lower()


def is_allowed_file(filename: str, allowed_extensions: list) -> bool:
    """检查文件扩展名是否允许"""
    ext = get_file_extension(filename)
    return ext in allowed_extensions


def format_file_size(size_bytes: int) -> str:
    """格式化文件大小"""
    for unit in ['B', 'KB', 'MB', 'GB']:
        if size_bytes < 1024.0:
            return f"{size_bytes:.2f} {unit}"
        size_bytes /= 1024.0
    return f"{size_bytes:.2f} TB"
