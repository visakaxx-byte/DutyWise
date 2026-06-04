"""
税率爬虫模块 - codeflagai.com
认证: AES-128-ECB加密登录 → CusAuthorization JWT → classification/search
"""
import asyncio
import base64
import json
import logging
from typing import Dict, Optional

import httpx
from Crypto.Cipher import AES
from Crypto.Util.Padding import pad

from app.core.config import settings

logger = logging.getLogger(__name__)

# AES-128 密钥（与前端 CryptoJS 一致）
_AES_KEY = b"imageBatchCompon"  # 16 bytes


def _aes_encrypt(plaintext: str) -> str:
    """AES-128-ECB 加密，PKCS7填充，输出 base64"""
    cipher = AES.new(_AES_KEY, AES.MODE_ECB)
    padded = pad(plaintext.encode("utf-8"), 16)
    return base64.b64encode(cipher.encrypt(padded)).decode("utf-8")


class TaxRateCrawler:
    """税率爬虫 — codeflagai.com 智能归类API"""

    BASE_HOST = "https://www.codeflagai.com"

    def __init__(self):
        self.timeout = settings.CRAWLER_TIMEOUT
        self._token: Optional[str] = None
        self._login_lock = asyncio.Lock()

    # ── 登录 ──────────────────────────────────────────────────

    async def _do_login(self) -> bool:
        """AES加密登录，从 CusAuthorization 响应头提取JWT"""
        username = settings.CRAWLER_USERNAME
        password = settings.CRAWLER_PASSWORD
        if not username or not password:
            logger.warning("爬虫账号未配置，跳过登录")
            return False

        try:
            login_payload = json.dumps({
                "loginname": username,
                "pwd": password,
                "phone": username,
                "code": "",
                "reqTime": "",
                "type": "",
            })
            encrypted = _aes_encrypt(login_payload)

            async with httpx.AsyncClient(timeout=self.timeout) as client:
                resp = await client.post(
                    f"{self.BASE_HOST}/xhqUser/login",
                    content=encrypted.encode(),
                    headers={"Content-Type": "text/plain"},
                )

                if resp.status_code != 200:
                    logger.warning(f"登录HTTP错误: {resp.status_code}")
                    self._token = None
                    return False

                body = resp.json()
                if body.get("code") != 200:
                    logger.warning(f"登录失败: {body.get('message', '未知错误')}")
                    self._token = None
                    return False

                # Token 在 CusAuthorization 响应头中
                auth_header = resp.headers.get("CusAuthorization", "")
                if not auth_header:
                    logger.warning("登录响应缺少 CusAuthorization 头")
                    self._token = None
                    return False

                self._token = auth_header
                logger.info("爬虫登录成功")
                return True

        except Exception as e:
            logger.error(f"爬虫登录异常: {e}")
            self._token = None
            return False

    async def _ensure_login(self, force: bool = False) -> bool:
        """确保已登录（带锁防并发），force=True 时强制重新登录"""
        async with self._login_lock:
            if self._token and not force:
                return True
            return await self._do_login()

    # ── API 查询 ──────────────────────────────────────────────

    async def _api_search(self, hs_code: str) -> Optional[Dict[str, Dict]]:
        """调用 classification/search 查询税率，返回 {hs_code: info} 字典"""
        headers = {
            "Content-Type": "application/json",
            "CusAuthorization": self._token,
        }
        payload = {
            "productNameCn": hs_code,
            "destinationCountryCode": "US",
            "destinationCountryName": "美国",
            "startCountryCode": "CN",
            "material": "",
            "isNew": True,
            "source": "classifySearch",
        }

        async with httpx.AsyncClient(timeout=self.timeout) as client:
            resp = await client.post(
                f"{self.BASE_HOST}/classification/search",
                json=payload,
                headers=headers,
            )

            if resp.status_code != 200:
                logger.warning(f"API HTTP {resp.status_code}: {hs_code}")
                return None

            try:
                body = resp.json()
            except json.JSONDecodeError:
                logger.error(f"API响应JSON解析失败: {hs_code}")
                return None

            if body.get("code") == 1401:
                logger.warning("Token已过期(1401)，清除缓存")
                self._token = None
                return None

            if body.get("code") != 200:
                logger.warning(f"API业务错误 {hs_code}: {body.get('message')}")
                return None

            return self._parse_classification_results(hs_code, body.get("data", {}))

    def _parse_classification_results(
        self, hs_code: str, data: dict
    ) -> Optional[Dict[str, Dict]]:
        """解析 classification/search 响应，提取所有备选编码及税率"""
        result_list = data.get("classificationResultList", [])
        code_list = data.get("classificationCodeList", [])

        if not result_list or not code_list:
            logger.info(f"未找到税率信息: {hs_code}")
            return None

        results: Dict[str, Dict] = {}
        for i in range(min(len(result_list), len(code_list))):
            result_entry = result_list[i]
            code_entry = code_list[i]

            matched_hs = str(result_entry.get("hsCode", "")).strip()
            if not matched_hs:
                continue

            tax_rate = code_entry.get("importTariffRate") or code_entry.get("columnRateOfDuty") or "N/A"
            description = result_entry.get("gName") or result_entry.get("dataWordCn") or ""

            anti_dumping_raw = code_entry.get("antiDumpingCountervailingRate")
            anti_dumping = bool(
                anti_dumping_raw
                and str(anti_dumping_raw).strip()
                and str(anti_dumping_raw).lower() not in ("none", "null", "")
            )

            results[matched_hs] = {
                "hs_code_cn": hs_code,
                "hs_code_us": matched_hs,
                "description_cn": description,
                "tax_rate": str(tax_rate),
                "anti_dumping": anti_dumping,
                "certification_required": False,
            }

        if not results:
            return None

        logger.info(f"HS编码 {hs_code} 找到 {len(results)} 个备选分类")
        return results

    async def search(self, hs_code: str) -> Optional[Dict[str, Dict]]:
        """
        查询单个HS编码的税率信息，返回 {匹配HS: info} 字典
        优先 API（遇登录过期自动重试一次），失败回退内置税率库
        """
        try:
            logger.info(f"查询HS编码: {hs_code}")

            await self._ensure_login()
            if self._token:
                result = await self._api_search(hs_code)
                if result:
                    return result

                # 如果 token 被清空（1401），重新登录后重试一次
                if self._token is None:
                    logger.info("重新登录后重试查询...")
                    if await self._ensure_login(force=True):
                        result = await self._api_search(hs_code)
                        if result:
                            return result

            builtin = self._get_builtin_rate(hs_code)
            if builtin:
                return {hs_code: builtin}
            return None

        except Exception as e:
            logger.warning(f"API查询失败 {hs_code}: {e}，使用内置税率库")
            builtin = self._get_builtin_rate(hs_code)
            if builtin:
                return {hs_code: builtin}
            return None

    # ── 内置税率库（API不可用时的回退） ───────────────────────

    def _get_builtin_rate(self, hs_code: str) -> Optional[Dict]:
        """内置美国进口税率参考库（按前缀长度优先匹配）"""
        RATE_DB = [
            ("8517900000", "电话设备零件", "0%", False),
            ("8517120000", "蜂窝网络电话", "0%", False),
            ("8517180000", "其他电话设备", "2.5%", False),
            ("8517120090", "智能手机", "2.5%", False),
            ("8517", "电话/通讯设备", "2.5%", False),
            ("8518300000", "蓝牙耳机", "4.9%", False),
            ("8518900000", "音频设备零件", "3.5%", False),
            ("8518", "音频设备/耳机", "4.9%", False),
            ("8471300000", "笔记本电脑", "0%", False),
            ("8471900000", "计算机外设", "0%", False),
            ("8471700000", "存储设备", "0%", False),
            ("8471", "计算机/数据处理", "0%", False),
            ("7326909000", "钢铁制品(其他)", "8.6%", True),
            ("7326901000", "钢铁制品(特定)", "5.1%", False),
            ("7326", "钢铁制品", "8.6%", True),
            ("3926909090", "塑料制品(其他)", "6.5%", False),
            ("3926", "塑料制品", "6.5%", False),
            ("6702909000", "人造花(其他)", "8.4%", False),
            ("6702", "人造花/装饰品", "8.4%", False),
        ]
        for prefix, desc, rate, ad in sorted(RATE_DB, key=lambda x: -len(x[0])):
            if hs_code.startswith(prefix):
                return {
                    "hs_code_cn": hs_code,
                    "hs_code_us": hs_code,
                    "description_cn": desc,
                    "tax_rate": rate,
                    "anti_dumping": ad,
                    "certification_required": False,
                }
        return {
            "hs_code_cn": hs_code,
            "hs_code_us": hs_code,
            "description_cn": "其他商品",
            "tax_rate": "3.5%",
            "anti_dumping": False,
            "certification_required": False,
        }

    # ── 批量查询 ──────────────────────────────────────────────

    async def batch_search(self, items: list) -> Dict[str, Dict]:
        """
        批量查询税率（每次调用前强制重新登录，确保token有效）
        """
        hs_codes = set()
        for item in items:
            code = item.get("商品编码")
            if code:
                hs_codes.add(str(code))

        if not hs_codes:
            logger.info("没有需要查询的HS编码")
            return {}

        logger.info(f"开始批量查询，共 {len(hs_codes)} 个HS编码")

        # 每次批量查询前强制重新登录
        self._token = None
        await self._ensure_login(force=True)

        # 并发查询（限制并发数，避免频控）
        semaphore = asyncio.Semaphore(1)

        async def search_one(code: str):
            async with semaphore:
                result = await self.search(code)
                await asyncio.sleep(settings.CRAWLER_DELAY)
                return code, result

        tasks = [search_one(code) for code in hs_codes]
        search_results = await asyncio.gather(*tasks, return_exceptions=True)

        results: Dict[str, Dict] = {}
        for item in search_results:
            if isinstance(item, Exception):
                logger.error(f"查询异常: {item}")
                continue
            code, result_dict = item
            if result_dict:
                results.update(result_dict)  # 合并所有备选编码

        logger.info(f"批量查询完成，共获得 {len(results)} 个编码")

        return self._enrich_with_candidates(results, hs_codes)

    def _enrich_with_candidates(
        self, results: Dict[str, Dict], hs_codes: set
    ) -> Dict[str, Dict]:
        """补充同前缀候选编码供优化器查找更优税率"""
        enriched = dict(results)
        CANDIDATE_DB = {
            "8517": {"8517900000", "8517120000", "8517180000"},
            "8518": {"8518900000", "8518100000"},
            "8471": {"8471900000", "8471700000"},
            "7326": {"7326901000", "7326200000"},
            "3926": {"3926901000"},
            "6702": {"6702901000"},
        }
        for hs_code in hs_codes:
            prefix4 = hs_code[:4]
            if prefix4 in CANDIDATE_DB:
                for cand_code in CANDIDATE_DB[prefix4]:
                    if cand_code not in enriched:
                        enriched[cand_code] = self._get_builtin_rate(cand_code)
        return enriched
