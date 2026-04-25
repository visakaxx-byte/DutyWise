"""
税率爬虫模块 - codeflagai.com
"""
import asyncio
import json
import logging
from typing import Dict, Optional
import httpx

from app.core.config import settings

logger = logging.getLogger(__name__)


class TaxRateCrawler:
    """税率爬虫"""

    def __init__(self):
        self.base_url = settings.CRAWLER_BASE_URL.rstrip("/")
        self.timeout = settings.CRAWLER_TIMEOUT
        self._token: Optional[str] = None
        self._login_lock = asyncio.Lock()

    async def _do_login(self) -> bool:
        """执行登录请求，返回是否成功"""
        username = settings.CRAWLER_USERNAME
        password = settings.CRAWLER_PASSWORD

        if not username or not password:
            logger.warning("爬虫账号未配置，跳过登录")
            return False

        try:
            login_url = f"{self.base_url}/api/user/login"
            async with httpx.AsyncClient(timeout=self.timeout) as client:
                resp = await client.post(login_url, json={
                    "username": username,
                    "password": password
                })
                data = resp.json()
                if data.get("code") == 200:
                    self._token = data.get("data", {}).get("token", "")
                    logger.info("爬虫登录成功")
                    return True
                else:
                    logger.warning(f"爬虫登录失败: {data.get('message', '未知错误')}")
                    self._token = None
                    return False
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

    async def _api_search(self, hs_code: str) -> Optional[Dict]:
        """纯API查询，遇到1401会自动清除token"""
        headers = {"Authorization": f"Bearer {self._token}"}
        async with httpx.AsyncClient(timeout=self.timeout) as client:
            response = await client.get(
                f"{self.base_url}/api/hscode/search",
                params={"code": hs_code},
                headers=headers
            )
            if response.status_code != 200:
                return None

            # 检查是否登录过期
            try:
                body = json.loads(response.text)
                if isinstance(body, dict) and body.get("code") == 1401:
                    logger.warning("Token已过期(1401)，清除缓存")
                    self._token = None
                    return None
            except json.JSONDecodeError:
                pass

            return self._parse_response(response.text, hs_code)

    async def search(self, hs_code: str) -> Optional[Dict]:
        """
        查询单个HS编码的税率信息
        优先 API（遇到登录过期自动重试一次），失败时使用内置税率库
        """
        try:
            logger.info(f"查询HS编码: {hs_code}")

            # 确保已登录
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

            # 回退到内置税率库
            return self._get_builtin_rate(hs_code)

        except Exception as e:
            logger.warning(f"API查询失败 {hs_code}: {e}，使用内置税率库")
            return self._get_builtin_rate(hs_code)

    def _get_builtin_rate(self, hs_code: str) -> Optional[Dict]:
        """内置美国进口税率参考库（按前缀长度优先匹配）"""
        RATE_DB = [
            # (前缀, 描述, 税率, 反倾销)
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
        # 按前缀长度降序匹配（优先精确匹配）
        for prefix, desc, rate, ad in sorted(RATE_DB, key=lambda x: -len(x[0])):
            if hs_code.startswith(prefix):
                return {
                    "hs_code_cn": hs_code,
                    "hs_code_us": hs_code,
                    "description_cn": desc,
                    "tax_rate": rate,
                    "anti_dumping": ad,
                    "certification_required": False
                }
        # 默认
        return {
            "hs_code_cn": hs_code,
            "hs_code_us": hs_code,
            "description_cn": "其他商品",
            "tax_rate": "3.5%",
            "anti_dumping": False,
            "certification_required": False
        }

    def _parse_response(self, text: str, hs_code: str) -> Optional[Dict]:
        """解析API响应"""
        try:
            data = json.loads(text)

            # codeflagai API 返回格式
            if isinstance(data, dict):
                if data.get("code") == 1401:
                    logger.warning(f"登录已过期: {data.get('message')}")
                    return None

                result = data.get("data") or data

                # 提取税率信息
                tax_rate = self._extract_tax_rate(result)
                description = result.get("goodsNameCn") or result.get("descriptionCn") or result.get("name", "")
                anti_dumping = self._check_anti_dumping(result)

                return {
                    "hs_code_cn": hs_code,
                    "hs_code_us": result.get("usHscode", hs_code),
                    "description_cn": description,
                    "tax_rate": tax_rate,
                    "anti_dumping": anti_dumping,
                    "certification_required": False
                }

        except (json.JSONDecodeError, Exception) as e:
            logger.error(f"解析响应失败 {hs_code}: {e}")

        return None

    def _extract_tax_rate(self, data: dict) -> str:
        """从返回数据中提取税率"""
        # 尝试多种可能的字段名
        for key in ("taxRate", "rate", "dutyRate", "tariffRate",
                     "generalRate", "mostFavoredNationRate",
                     "general_rate", "mfn_rate"):
            val = data.get(key)
            if val is not None and val != "":
                return str(val)

        # 如果所有字段都没有，尝试嵌套结构
        tariff = data.get("tariff", {})
        if isinstance(tariff, dict):
            for key in ("general", "mfn", "rate"):
                val = tariff.get(key)
                if val is not None and val != "":
                    return str(val)

        return "N/A"

    def _check_anti_dumping(self, data: dict) -> bool:
        """检查是否有反倾销标记"""
        for key in ("antiDumping", "anti_dumping", "isAntiDumping"):
            val = data.get(key)
            if val is True or str(val).lower() in ("true", "yes", "1"):
                return True

        # 检查描述中是否包含反倾销关键词
        desc = str(data.get("descriptionCn", "")) + str(data.get("remark", ""))
        if "反倾销" in desc or "anti-dumping" in desc.lower():
            return True

        return False

    async def batch_search(self, items: list) -> Dict[str, Dict]:
        """
        批量查询税率（每次调用前强制重新登录，确保token有效）
        """
        results = {}

        # 提取所有HS编码
        hs_codes = set()
        for item in items:
            hs_code = item.get("商品编码")
            if hs_code:
                hs_codes.add(str(hs_code))

        if not hs_codes:
            logger.info("没有需要查询的HS编码")
            return results

        logger.info(f"开始批量查询，共 {len(hs_codes)} 个HS编码")

        # 每次批量查询前强制重新登录，确保session有效
        self._token = None
        await self._ensure_login(force=True)

        # 并发查询（限制并发数）
        semaphore = asyncio.Semaphore(3)

        async def search_with_semaphore(code):
            async with semaphore:
                result = await self.search(code)
                await asyncio.sleep(settings.CRAWLER_DELAY)
                return code, result

        tasks = [search_with_semaphore(code) for code in hs_codes]
        search_results = await asyncio.gather(*tasks, return_exceptions=True)

        for item in search_results:
            if isinstance(item, Exception):
                logger.error(f"查询异常: {item}")
                continue
            code, result = item
            if result:
                results[code] = result

        logger.info(f"批量查询完成，成功 {len(results)}/{len(hs_codes)} 个")

        # 补充同前缀候选编码供优化器对比
        return self._enrich_with_candidates(results, hs_codes)

    def _enrich_with_candidates(self, results: Dict[str, Dict], hs_codes: set) -> Dict[str, Dict]:
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
