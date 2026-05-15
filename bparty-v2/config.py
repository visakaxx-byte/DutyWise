from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

from dotenv import dotenv_values


APP_DIR = Path(__file__).resolve().parent
LOCAL_ENV = APP_DIR / ".env"


def env_value(*names: str, default: str = "") -> str:
    env_file = dotenv_values(LOCAL_ENV) if LOCAL_ENV.exists() else {}
    for name in names:
        value = os.getenv(name)
        if value:
            return value
        file_value = env_file.get(name)
        if file_value:
            return str(file_value)
    return default


@dataclass(frozen=True)
class LLMSettings:
    api_key: str
    base_url: str
    model: str
    timeout: float = 90.0


@dataclass(frozen=True)
class CrawlerSettings:
    username: str
    password: str
    base_url: str
    timeout: float = 30.0
    delay: float = 0.4
    max_retries: int = 4
    rate_limit_backoff: float = 8.0


def get_llm_settings() -> LLMSettings:
    return LLMSettings(
        api_key=env_value("LLM_API_KEY", "DOUBAO_API_KEY"),
        base_url=env_value("LLM_BASE_URL", "DOUBAO_ENDPOINT", "DOUBAO_BASE_URL", default="https://ark.cn-beijing.volces.com/api/coding/v3").rstrip("/"),
        model=env_value("LLM_MODEL", "DOUBAO_MODEL", default="doubao-pro-32k"),
        timeout=float(env_value("LLM_TIMEOUT", default="240") or 240),
    )


def get_crawler_settings() -> CrawlerSettings:
    return CrawlerSettings(
        username=env_value("CRAWLER_USERNAME"),
        password=env_value("CRAWLER_PASSWORD"),
        base_url=env_value("CRAWLER_BASE_URL", default="https://www.codeflagai.com").rstrip("/").replace("/index", ""),
        timeout=float(env_value("CRAWLER_TIMEOUT", default="30") or 30),
        delay=float(env_value("CRAWLER_DELAY", default="3") or 3),
        max_retries=int(env_value("CRAWLER_MAX_RETRIES", default="4") or 4),
        rate_limit_backoff=float(env_value("CRAWLER_RATE_LIMIT_BACKOFF", default="8") or 8),
    )


def require_runtime_config() -> dict:
    llm = get_llm_settings()
    crawler = get_crawler_settings()
    missing = []
    if not llm.api_key:
        missing.append("LLM_API_KEY 或 DOUBAO_API_KEY")
    if not llm.model:
        missing.append("LLM_MODEL 或 DOUBAO_MODEL")
    if not crawler.username:
        missing.append("CRAWLER_USERNAME")
    if not crawler.password:
        missing.append("CRAWLER_PASSWORD")
    if missing:
        raise RuntimeError("缺少必需配置，不能执行 LLM/爬虫流程: " + ", ".join(missing))

    return {
        "llm": {"base_url": llm.base_url, "model": llm.model},
        "crawler": {"base_url": crawler.base_url, "username_configured": True},
    }
