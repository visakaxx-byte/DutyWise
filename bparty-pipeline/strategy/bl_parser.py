"""
提单品名提取模块

从 PDF / 文本提单中提取商品名称列表。
核心策略: 提单头部的大写英文短语即为品名，过滤地址/物流/数字信息。

海运提单示例输出: ["SILICONE COASTER", "PLASTIC MOBILE PHONE STAND"]
空运提单示例输出: ["CONSOL"] (拼箱货，品名详见发票)
"""

import logging
import re
from pathlib import Path
from typing import List, Set

logger = logging.getLogger(__name__)


class BLParser:
    """提单品名解析器"""

    # ── 噪音模式 ──
    NOISE_PATTERNS = [
        r'(?i)\b(LLC|INC|LTD|CORP|CO\.|LIMITED|CORPORATION|'
        r'IMPORT|EXPORT|SUPPLY|CHAIN|MANAGEMENT|LOGISTICS|'
        r'FORWARDER|CARRIER|SHIPPING|EXPRESS|AIRLINES|'
        r'FREIGHT|PREPAID|COLLECT|BROKER|AGENT|CUSTOMS)\b',
        r'(?i)\b(ROAD|STREET|AVE|AVENUE|BLVD|PKWY|LANE|DRIVE|'
        r'BUILDING|FLOOR|SUITE|ROOM|DISTRICT|PROVINCE|'
        r'BOULEVARD|HIGHWAY|PLACE|COURT|CIRCLE|'
        r'SHENZHEN|SHANGHAI|NINGBO|GUANGZHOU|YIWU|'
        r'QINGDAO|TIANJIN|XIAMEN|HONG\s?KONG|SICHUAN|CHENGDU|'
        r'BEACH|LOS\s?ANGELES|NEW\s?YORK|VIRGINIA|CHESAPEAKE|ONTARIO)\b',
        r'(?i)(?:TEL|FAX|EMAIL|PHONE|BROKER)\s*[:@]',
        r'^\d{2,4}[-/]\d{2}[-/]\d{2,4}',
        r'^[A-Z]{4}\d{6,7}',
        r'^\d+\.?\d*\s*(KGS?|CBM|CFT|CTNS?|CARTONS?|PKGS?)\s*$',
        r'(?i)(40|20)(?:GP|HQ|RF|OT|FR|HC|DC)\s*(FCL|LCL)?',
        r'^[A-Z]{3,4}\d{8,12}',
        r'(?:上海|深圳|广州|宁波|天津|青岛|厦门|香港|四川|成都|洛杉矶|纽约)',
        r'^[\d\s,.\-/()+:\'\\]+$',
        r'[A-Z]{4}\s*\d{6,7}',
        r'(?i)(FRIMS?|CODE|TEL|ATTN|FAX)\b',
        r'@\w+\.\w+',
    ]

    # 产品特征词 (候选品名必须含至少一个)
    PRODUCT_FEATURES: Set[str] = {
        "coaster", "phone", "stand", "plastic", "silicone",
        "bag", "box", "case", "cover", "cup", "pad", "mat",
        "toy", "tool", "light", "lamp", "chair", "table",
        "rack", "shelf", "holder", "bracket", "hook",
        "wire", "cable", "charger", "adapter", "plug",
        "clothes", "shirt", "shoes", "hat", "cap", "glove",
        "glass", "metal", "wood", "paper", "fabric",
        "bottle", "jar", "container", "tank",
        "tape", "film", "sheet", "roll", "sticker",
        "towel", "blanket", "pillow", "curtain",
        "flower", "plant", "decor", "ornament",
        "ball", "game", "puzzle", "racket", "paddle",
        "fan", "heater", "cooler", "pump", "valve",
        "bike", "scooter", "helmet", "lock",
        "speaker", "camera", "monitor", "screen",
        "food", "candy", "snack", "drink",
        "cosmetics", "cream", "oil", "soap",
        "gift", "card", "book", "calendar", "notebook",
        "board", "panel", "display", "frame",
        "pole", "rod", "pipe", "tube", "hose",
        "motor", "engine", "machine", "equipment",
        "fitting", "connector", "valve", "switch",
        "filter", "cartridge",
        "fence", "net", "mesh", "screen",
        "tape", "glue", "adhesive", "sealant",
        "cleaner", "detergent", "polish", "wax",
        "patch", "repair", "kit", "set",
    }

    RE_DIGIT = re.compile(r'(?<=[A-Za-z])(?=\d)')
    RE_UNIT = re.compile(r'(?<=\d)(?=[A-Z])')

    def parse_pdf(self, pdf_path: str) -> List[str]:
        """从PDF提单中提取品名列表"""
        if not Path(pdf_path).exists():
            logger.error(f"提单文件不存在: {pdf_path}")
            return []

        text = self._extract_pdf_text(pdf_path)
        if not text:
            return []
        return self.parse_text(text)

    def parse_text(self, text: str) -> List[str]:
        """从提单文本中提取品名列表"""
        text = text.replace("\r\n", "\n").replace("\r", "\n")
        lines = text.split("\n")

        all_candidates: List[str] = []
        for line in lines:
            line = line.strip()
            if not line:
                continue

            # 拆分字母→数字粘连
            line = self.RE_DIGIT.sub("|", line)
            line = self.RE_UNIT.sub("|", line)

            fragments = [f.strip() for f in line.split("|") if f.strip()]
            for frag in fragments:
                sub = [s.strip() for s in re.split(r'[,;]{2,}|\s{3,}', frag) if s.strip()]
                for s in sub:
                    if len(s) >= 3:
                        all_candidates.append(s)

        # 过滤 + 品名校验
        products: List[str] = []
        seen = set()
        for c in all_candidates:
            cleaned = self._clean(c)
            if not cleaned:
                continue
            if not self._is_product_name(cleaned):
                continue
            key = cleaned.lower()
            if key not in seen:
                seen.add(key)
                products.append(cleaned)

        # 如果严格模式没结果，放松一次（不要求产品特征词）
        if not products:
            for c in all_candidates:
                cleaned = self._clean(c)
                if not cleaned:
                    continue
                key = cleaned.lower()
                if key not in seen:
                    seen.add(key)
                    products.append(cleaned)

        logger.info(f"提单品名: {len(products)} → {products}")
        return products

    def _clean(self, text: str) -> str:
        """清洗单个候选"""
        t = text.strip().rstrip(".,;:\"'")

        if len(t) < 3 or len(t) > 80:
            return ""
        if re.match(r'^[\d\s,./\-:;()]+$', t):
            return ""

        for pat in self.NOISE_PATTERNS:
            if re.search(pat, t):
                return ""

        # 去尾部单位
        t = re.sub(r'\s+(CBM|KGS?|CTNS?|CARTONS?|PCS|SETS?|PKGS?)$', '', t, flags=re.IGNORECASE).strip()
        return t

    def _is_product_name(self, text: str) -> bool:
        """
        判断文本是否像品名。

        规则:
          - 必须含大写字母
          - 必须含 >=1 个产品特征词
          - 不能以数字开头
          - 非纯代码/缩写
        """
        if not text:
            return False

        # 至少含一个产品特征词
        words = set(text.lower().replace("-", " ").replace("/", " ").split())
        if not words & self.PRODUCT_FEATURES:
            return False

        # 不能以数字开头
        if text[0].isdigit():
            return False

        # 至少含一个大写字母 (产品名通常大写)
        if not any(c.isupper() for c in text):
            return False

        return True

    # ── PDF 文本提取 ──

    def _extract_pdf_text(self, pdf_path: str) -> str:
        """从PDF文件中提取纯文本"""
        try:
            from PyPDF2 import PdfReader
        except ImportError:
            logger.error("PyPDF2 未安装")
            return ""

        try:
            reader = PdfReader(pdf_path)
            return "\n".join(p.extract_text() or "" for p in reader.pages)
        except Exception as e:
            logger.exception(f"PDF解析失败: {e}")
            return ""
