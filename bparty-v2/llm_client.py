from __future__ import annotations

import asyncio
import json
import re
from typing import Any, Optional

import httpx

from config import LLMSettings, get_llm_settings


class LLMError(RuntimeError):
    pass


class LLMClient:
    def __init__(self, settings: Optional[LLMSettings] = None):
        self.settings = settings or get_llm_settings()
        if not self.settings.api_key:
            raise LLMError("缺少 LLM_API_KEY 或 DOUBAO_API_KEY")

    async def chat_json(
        self,
        messages: list[dict[str, Any]],
        *,
        temperature: float = 0.1,
        model: Optional[str] = None,
        max_tokens: int = 4096,
        json_mode: bool = True,
    ) -> dict[str, Any]:
        payload = {
            "model": model or self.settings.model,
            "messages": messages,
            "temperature": temperature,
            "max_tokens": max_tokens,
        }
        if json_mode:
            payload["response_format"] = {"type": "json_object"}
        return await self._post_chat_json(payload)

    async def chat_json_with_images(
        self,
        text: str,
        image_data_urls: list[str],
        *,
        temperature: float = 0.0,
        model: Optional[str] = None,
        max_tokens: int = 4096,
    ) -> dict[str, Any]:
        content: list[dict[str, Any]] = [{"type": "text", "text": text}]
        content.extend({"type": "image_url", "image_url": {"url": url}} for url in image_data_urls)
        payload = {
            "model": model or self.settings.vision_model or self.settings.model,
            "messages": [{"role": "user", "content": content}],
            "temperature": temperature,
            "max_tokens": max_tokens,
        }
        return await self._post_chat_json(payload)

    async def _post_chat_json(self, payload: dict[str, Any]) -> dict[str, Any]:
        headers = {
            "Authorization": f"Bearer {self.settings.api_key}",
            "Content-Type": "application/json",
        }
        timeout = httpx.Timeout(connect=20.0, read=self.settings.timeout, write=30.0, pool=20.0)
        response = await self._post_chat_completion(payload, headers, timeout)
        if response.status_code != 200 and should_retry_without_json_mode(response, payload):
            fallback_payload = dict(payload)
            fallback_payload.pop("response_format", None)
            response = await self._post_chat_completion(fallback_payload, headers, timeout)

        if response.status_code != 200:
            raise LLMError(f"LLM HTTP {response.status_code}: {response.text[:500]}")

        try:
            body = response.json()
            content = body["choices"][0]["message"]["content"]
        except Exception as exc:
            raise LLMError(f"LLM 响应结构异常: {response.text[:500]}") from exc

        return parse_json_object(content)

    async def _post_chat_completion(
        self,
        payload: dict[str, Any],
        headers: dict[str, str],
        timeout: httpx.Timeout,
    ) -> httpx.Response:
        last_timeout = None
        for attempt in range(1, 3):
            try:
                async with httpx.AsyncClient(timeout=timeout) as client:
                    response = await client.post(f"{self.settings.base_url}/chat/completions", headers=headers, json=payload)
                break
            except httpx.TimeoutException as exc:
                last_timeout = exc
                if attempt >= 2:
                    raise LLMError(f"LLM 请求超时，已重试 {attempt} 次") from exc
                await asyncio.sleep(3)
        else:
            raise LLMError("LLM 请求超时") from last_timeout
        return response


def should_retry_without_json_mode(response: httpx.Response, payload: dict[str, Any]) -> bool:
    if "response_format" not in payload or response.status_code not in {400, 404, 422}:
        return False
    text = response.text.lower()
    return "response_format" in text or "json_object" in text


def parse_json_object(content: str) -> dict[str, Any]:
    try:
        parsed = json.loads(content)
        if isinstance(parsed, dict):
            return parsed
    except json.JSONDecodeError:
        pass

    match = re.search(r"\{.*\}", content, flags=re.DOTALL)
    if not match:
        raise LLMError("LLM 未返回 JSON object")
    try:
        parsed = json.loads(match.group(0))
    except json.JSONDecodeError as exc:
        raise LLMError("LLM 返回 JSON 无法解析") from exc
    if not isinstance(parsed, dict):
        raise LLMError("LLM JSON 顶层必须是 object")
    return parsed
