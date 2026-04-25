"""
配置文件
"""
from pydantic_settings import BaseSettings
from typing import Optional


class Settings(BaseSettings):
    """应用配置"""

    # 应用配置
    APP_NAME: str = "清关文件处理系统"
    APP_VERSION: str = "1.0.0"
    DEBUG: bool = True

    # 服务器配置
    HOST: str = "0.0.0.0"
    PORT: int = 8000

    # CORS 配置
    CORS_ORIGINS: list = ["http://localhost:3000", "http://localhost:5173"]

    # 文件上传配置
    UPLOAD_DIR: str = "./uploads"
    OUTPUT_DIR: str = "./outputs"
    MAX_FILE_SIZE: int = 50 * 1024 * 1024  # 50MB
    ALLOWED_EXTENSIONS: list = [".xlsx", ".xls"]

    # Redis 配置
    REDIS_HOST: str = "localhost"
    REDIS_PORT: int = 6379
    REDIS_DB: int = 0
    REDIS_PASSWORD: Optional[str] = None

    # Celery 配置
    CELERY_BROKER_URL: str = "redis://localhost:6379/0"
    CELERY_RESULT_BACKEND: str = "redis://localhost:6379/0"

    # 豆包 API 配置
    DOUBAO_API_KEY: Optional[str] = None
    DOUBAO_ENDPOINT: str = "https://ark.cn-beijing.volces.com/api/v3"
    DOUBAO_MODEL: str = "doubao-pro-32k"

    # 爬虫配置
    CRAWLER_TIMEOUT: int = 30
    CRAWLER_RETRY: int = 3
    CRAWLER_DELAY: float = 1.0

    # 缓存配置
    CACHE_TTL: int = 7 * 24 * 3600  # 7天

    class Config:
        env_file = ".env"
        case_sensitive = True


settings = Settings()
