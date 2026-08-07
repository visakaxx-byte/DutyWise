from __future__ import annotations

import asyncio
import base64
import json
from typing import Any, Optional

import httpx
from Crypto.Cipher import AES
from Crypto.Util.Padding import pad, unpad

from config import CrawlerSettings, get_crawler_settings


AES_KEY = b"imageBatchCompon"
AUTH_EXPIRED_MARKERS = (
    "你已下线",
    "重新登陆",
    "重新登录",
    "请登录",
    "登录失效",
    "登陆失效",
    "另一地点登录",
    "token",
)


class CrawlerError(RuntimeError):
    pass


def aes_encrypt(plaintext: str) -> str:
    cipher = AES.new(AES_KEY, AES.MODE_ECB)
    return base64.b64encode(cipher.encrypt(pad(plaintext.encode("utf-8"), 16))).decode("utf-8")


def aes_decrypt(ciphertext: str) -> str:
    cipher = AES.new(AES_KEY, AES.MODE_ECB)
    raw = base64.b64decode(ciphertext)
    return unpad(cipher.decrypt(raw), 16).decode("utf-8")


def parse_encrypted_login_data(value: Any) -> dict[str, Any] | None:
    if not isinstance(value, str) or not value.strip():
        return None
    try:
        decrypted = aes_decrypt(value.strip())
        payload = json.loads(decrypted)
    except (ValueError, json.JSONDecodeError):
        return None
    return payload if isinstance(payload, dict) else None


def login_body_error_detail(body: dict[str, Any]) -> str:
    data_payload = parse_encrypted_login_data(body.get("data"))
    if data_payload:
        message = data_payload.get("message") or data_payload.get("msg")
        if message:
            return str(message)
        code = data_payload.get("code")
        if code not in (None, 200, "200"):
            return str(code)
    message = body.get("message") or body.get("msg")
    if message and body.get("code") != 200:
        return str(message)
    return "缺少 CusAuthorization"


def response_error_detail(body: dict[str, Any]) -> str:
    return str(body.get("message") or body.get("msg") or body.get("code") or "")


def is_auth_expired(message: str) -> bool:
    lowered = str(message or "").lower()
    return any(marker in message or marker in lowered for marker in AUTH_EXPIRED_MARKERS)


class StrictTaxCrawler:
    def __init__(self, settings: Optional[CrawlerSettings] = None):
        self.settings = settings or get_crawler_settings()
        if not self.settings.username or not self.settings.password:
            raise CrawlerError("缺少 CRAWLER_USERNAME 或 CRAWLER_PASSWORD")
        self._token: Optional[str] = None
        self.auth_expired_retry_count = 0

    async def login(self) -> None:
        login_payload = json.dumps(
            {
                "loginname": self.settings.username,
                "pwd": self.settings.password,
                "phone": self.settings.username,
                "code": "",
                "reqTime": "",
                "type": "",
            },
            ensure_ascii=False,
        )
        encrypted = aes_encrypt(login_payload)
        async with httpx.AsyncClient(timeout=self.settings.timeout) as client:
            response = await client.post(
                f"{self.settings.base_url}/xhqUser/login",
                content=encrypted.encode("utf-8"),
                headers={"Content-Type": "text/plain"},
            )

        if response.status_code != 200:
            raise CrawlerError(f"爬虫登录 HTTP {response.status_code}")
        try:
            body = response.json()
        except json.JSONDecodeError as exc:
            raise CrawlerError("爬虫登录响应不是 JSON") from exc
        if body.get("code") != 200:
            raise CrawlerError(f"爬虫登录失败: {body.get('message') or body.get('msg') or body.get('code')}")
        token = response.headers.get("CusAuthorization")
        if not token:
            raise CrawlerError(f"爬虫登录失败: {login_body_error_detail(body)}")
        self._token = token

    async def search(self, hs_code: str) -> dict[str, dict[str, Any]]:
        if not self._token:
            await self.login()
        assert self._token

        last_error = ""
        for attempt in range(1, self.settings.max_retries + 1):
            try:
                return await self._search_once(hs_code)
            except CrawlerError as exc:
                last_error = str(exc)
                if not is_rate_limited(last_error) or attempt >= self.settings.max_retries:
                    raise
                await asyncio.sleep(self.settings.rate_limit_backoff * attempt)
        raise CrawlerError(last_error or f"HS {hs_code} 查询失败")

    async def _search_once(self, hs_code: str, *, auth_retry: bool = True) -> dict[str, dict[str, Any]]:
        assert self._token
        payload = {
            "productNameCn": hs_code,
            "destinationCountryCode": "US",
            "destinationCountryName": "美国",
            "startCountryCode": "CN",
            "material": "",
            "isNew": True,
            "source": "classifySearch",
        }
        async with httpx.AsyncClient(timeout=self.settings.timeout) as client:
            response = await client.post(
                f"{self.settings.base_url}/classification/search",
                headers={"Content-Type": "application/json", "CusAuthorization": self._token},
                json=payload,
            )

        if response.status_code != 200:
            raise CrawlerError(f"HS {hs_code} 查询 HTTP {response.status_code}")
        try:
            body = response.json()
        except json.JSONDecodeError as exc:
            raise CrawlerError(f"HS {hs_code} 查询响应不是 JSON") from exc

        error_detail = response_error_detail(body)
        if (body.get("code") == 1401 or is_auth_expired(error_detail)) and auth_retry:
            self.auth_expired_retry_count += 1
            self._token = None
            await self.login()
            return await self._search_once(hs_code, auth_retry=False)
        if body.get("code") != 200:
            raise CrawlerError(f"HS {hs_code} 查询失败: {error_detail}")

        parsed = parse_classification_results(hs_code, body.get("data") or {})
        if not parsed:
            raise CrawlerError(f"HS {hs_code} 未获得税率结果")
        return parsed

    async def search_product(self, product_name: str, material: str = "") -> dict[str, dict[str, Any]]:
        if not self._token:
            await self.login()
        assert self._token

        last_error = ""
        for attempt in range(1, self.settings.max_retries + 1):
            try:
                return await self._search_product_once(product_name, material)
            except CrawlerError as exc:
                last_error = str(exc)
                if not is_rate_limited(last_error) or attempt >= self.settings.max_retries:
                    raise
                await asyncio.sleep(self.settings.rate_limit_backoff * attempt)
        raise CrawlerError(last_error or f"商品 {product_name} 查询失败")

    async def _search_product_once(
        self,
        product_name: str,
        material: str = "",
        *,
        auth_retry: bool = True,
    ) -> dict[str, dict[str, Any]]:
        assert self._token
        payload = {
            "productNameCn": product_name,
            "destinationCountryCode": "US",
            "destinationCountryName": "美国",
            "startCountryCode": "CN",
            "material": material or "",
            "isNew": True,
            "source": "classifySearch",
        }
        async with httpx.AsyncClient(timeout=self.settings.timeout) as client:
            response = await client.post(
                f"{self.settings.base_url}/classification/search",
                headers={"Content-Type": "application/json", "CusAuthorization": self._token},
                json=payload,
            )

        if response.status_code != 200:
            raise CrawlerError(f"商品 {product_name} 查询 HTTP {response.status_code}")
        try:
            body = response.json()
        except json.JSONDecodeError as exc:
            raise CrawlerError(f"商品 {product_name} 查询响应不是 JSON") from exc

        error_detail = response_error_detail(body)
        if (body.get("code") == 1401 or is_auth_expired(error_detail)) and auth_retry:
            self.auth_expired_retry_count += 1
            self._token = None
            await self.login()
            return await self._search_product_once(product_name, material, auth_retry=False)
        if body.get("code") != 200:
            raise CrawlerError(f"商品 {product_name} 查询失败: {error_detail}")

        parsed = parse_classification_results(product_name, body.get("data") or {})
        if not parsed:
            raise CrawlerError(f"商品 {product_name} 未获得智能归类结果")
        return parsed

    async def batch_search(self, hs_codes: list[str]) -> dict[str, dict[str, Any]]:
        unique_codes = []
        seen = set()
        for code in hs_codes:
            normalized = normalize_hs(code)
            if normalized and normalized not in seen:
                seen.add(normalized)
                unique_codes.append(normalized)
        if not unique_codes:
            raise CrawlerError("没有可查询的 HS 编码")

        await self.login()
        results: dict[str, dict[str, Any]] = {}
        for index, code in enumerate(unique_codes):
            if index > 0:
                await asyncio.sleep(self.settings.delay)
            found = await self.search(code)
            results.update(found)
        return results


def parse_classification_results(query: str, data: dict[str, Any]) -> dict[str, dict[str, Any]]:
    result_list = data.get("classificationResultList") or []
    code_list = data.get("classificationCodeList") or []
    parsed: dict[str, dict[str, Any]] = {}
    for index, code_entry in enumerate(code_list):
        if not isinstance(code_entry, dict):
            continue
        result_entry = result_list[index] if index < len(result_list) and isinstance(result_list[index], dict) else {}
        matched_hs = normalize_hs(code_entry.get("hsCode"))
        if not matched_hs:
            continue
        tax_rate = code_entry.get("importTariffRate") or code_entry.get("columnRateOfDuty") or "N/A"
        additional_tax_rate = (
            code_entry.get("additionalTaxRate")
            or code_entry.get("additionalTariffRate")
            or code_entry.get("additionalTariff")
            or code_entry.get("addTariffRate")
            or code_entry.get("extraTariffRate")
            or ""
        )
        additional_tax_details = (
            code_entry.get("additionalTaxList")
            or code_entry.get("additionalTariffList")
            or code_entry.get("additionalTaxRateList")
            or code_entry.get("additionalTariffRateList")
            or code_entry.get("extraTariffList")
            or []
        )
        anti_dumping_raw = code_entry.get("antiDumpingCountervailingRate")
        certification_texts = collect_certification_texts(code_entry) + collect_certification_texts(result_entry)
        parsed[matched_hs] = {
            "query": str(query or ""),
            "hs_code_cn": normalize_hs(result_entry.get("hsCode") or query),
            "hs_code_us": matched_hs,
            "description_cn": (
                code_entry.get("taricCn")
                or code_entry.get("taric")
                or result_entry.get("gName")
                or result_entry.get("dataWordCn")
                or ""
            ),
            "source_description_cn": result_entry.get("gName") or result_entry.get("dataWordCn") or "",
            "taric": code_entry.get("taric") or "",
            "tax_rate": str(tax_rate),
            "additional_tax_rate": str(additional_tax_rate),
            "additional_tax_details": additional_tax_details,
            "applicable_tax_rate": str(code_entry.get("applicableTaxRate") or ""),
            "column_rate_of_duty": str(code_entry.get("columnRateOfDuty") or ""),
            "anti_dumping": bool(
                anti_dumping_raw
                and str(anti_dumping_raw).strip()
                and str(anti_dumping_raw).lower() not in ("none", "null", "")
            ),
            "anti_dumping_rate": str(anti_dumping_raw or ""),
            "certification_required": bool(certification_texts),
            "certification_texts": certification_texts,
        }
    return parsed


def collect_certification_texts(value: Any) -> list[str]:
    key_markers = (
        "cert",
        "认证",
        "监管",
        "authentication",
        "license",
        "licence",
        "permit",
        "regulat",
        "supervision",
        "qualification",
        "document",
        "fda",
        "fcc",
        "cpsc",
        "cpsia",
        "tsca",
        "lacey",
    )
    found: list[str] = []

    def walk(item: Any, key_hint: str = "") -> None:
        if isinstance(item, dict):
            for key, inner in item.items():
                walk(inner, str(key))
            return
        if isinstance(item, list):
            for inner in item:
                walk(inner, key_hint)
            return
        if item in (None, "", False):
            return

        text = str(item).strip()
        key_text = key_hint.lower()
        text_lower = text.lower()
        if any(marker in key_text or marker in text_lower for marker in key_markers):
            if text not in found:
                found.append(text)

    walk(value)
    return found


def normalize_hs(value: Any) -> str:
    return "".join(ch for ch in str(value or "") if ch.isdigit())


def is_rate_limited(message: str) -> bool:
    lowered = message.lower()
    return "频繁" in message or "rate" in lowered or "too many" in lowered or "稍候" in message
