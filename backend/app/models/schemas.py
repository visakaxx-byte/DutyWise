"""
数据模型
"""
from pydantic import BaseModel
from typing import List, Optional, Dict, Any
from datetime import datetime
from enum import Enum


class TaskStatus(str, Enum):
    """任务状态"""
    PENDING = "pending"
    PROCESSING = "processing"
    COMPLETED = "completed"
    FAILED = "failed"


class RiskLevel(str, Enum):
    """风险等级"""
    LOW = "LOW"
    MEDIUM = "MEDIUM"
    HIGH = "HIGH"


class UploadFileInfo(BaseModel):
    """上传文件信息"""
    filename: str
    size: int
    path: str


class ProcessOptions(BaseModel):
    """处理选项"""
    exclude_anti_dumping: bool = False
    min_similarity: float = 0.15


class TaskProgress(BaseModel):
    """任务进度"""
    task_id: str
    status: TaskStatus
    progress: int
    current_step: str
    error_message: Optional[str] = None


class ShipmentItem(BaseModel):
    """票据商品项"""
    item_id: int
    original_name_cn: str
    original_hs_code: Optional[str] = None
    original_tax_rate: Optional[str] = None
    optimized_hs_code: Optional[str] = None
    optimized_tax_rate: Optional[str] = None
    is_optimized: bool = False
    similarity_score: Optional[float] = None
    risk_level: RiskLevel = RiskLevel.LOW
    risk_warnings: List[str] = []


class ShipmentStatistics(BaseModel):
    """票据统计信息"""
    total_items: int
    optimized_items: int
    optimization_rate: float
    high_risk_items: int = 0
    conflict_items: int = 0


class ShipmentResult(BaseModel):
    """票据处理结果"""
    shipment_id: int
    shipment_no: str
    status: TaskStatus
    statistics: ShipmentStatistics
    files: Dict[str, str]
    warnings: List[Dict[str, Any]] = []


class APIResponse(BaseModel):
    """统一响应格式"""
    code: int
    message: str
    data: Optional[Any] = None
