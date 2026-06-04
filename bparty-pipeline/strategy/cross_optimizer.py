"""
跨章HS编码优化模块

与原 optimizer.py 的核心区别:
  - 不再限制同4位前缀: 全量搜索 tax_data 中所有编码
  - 章节偏好加权: preferred_chapters 白名单候选获得额外相似度加成
  - 可调权重: tax_weight / similarity_weight 从 config 读取
  - 可调阈值: adopt_threshold 从 config 读取
"""

import logging
from typing import Dict, List, Any, Optional

logger = logging.getLogger(__name__)


class CrossChapterOptimizer:
    """跨章HS编码优化器: 全量搜索 + 章节偏好加权"""

    def __init__(self, config: Dict[str, Any] = None):
        cfg = config or {}
        self.min_similarity = cfg.get("min_similarity", 0.10)
        self.tax_weight = cfg.get("tax_weight", 0.7)
        self.similarity_weight = cfg.get("similarity_weight", 0.3)
        self.adopt_threshold = cfg.get("adopt_threshold", 0.50)
        self.preferred_chapters = set(cfg.get("preferred_chapters", []))
        self.chapter_boost = cfg.get("chapter_boost", 0.15)
        self.exclude_anti_dumping = cfg.get("exclude_anti_dumping", False)
        self.exclude_cert = cfg.get("exclude_certification_required", True)

    def optimize(
        self,
        items: List[Dict[str, Any]],
        tax_data: Dict[str, Dict],
    ) -> Dict[str, Any]:
        """
        批量跨章HS优化。

        Args:
            items: 商品列表 (含 商品编码、中文品名)
            tax_data: 税率数据 {hs_code: {tax_rate, anti_dumping, ...}}

        Returns:
            {"items": [...], "optimization_logs": [...]}
        """
        optimized_items = []
        optimization_logs = []

        for item in items:
            # 提单保护项不参与跨章HS优化
            if item.get("_bl_protected"):
                optimized_items.append(item)
                continue

            original_hs = str(item.get("商品编码", ""))
            original_name = str(item.get("中文品名", ""))

            # 获取原始税率
            original_info = tax_data.get(original_hs)
            original_rate = original_info.get("tax_rate", "") if original_info else ""

            # 查找最佳候选
            best = self._find_best_candidate(original_hs, original_name, tax_data)

            if best and best["is_better"]:
                optimized_item = item.copy()
                optimized_item["商品编码"] = best["hs_code"]
                optimized_item["_original_hs_code"] = original_hs
                optimized_item["_optimized"] = True
                optimized_item["_similarity_score"] = best["similarity"]
                optimized_items.append(optimized_item)

                log = {
                    "original_name": original_name,
                    "original_hs_code": original_hs,
                    "original_tax_rate": original_rate,
                    "new_hs_code": best["hs_code"],
                    "new_tax_rate": best["tax_rate"],
                    "reason": f"跨章优化 (score={best['final_score']:.3f})",
                    "similarity_score": best["similarity"],
                    "risk_warning": (
                        "反倾销标记" if best.get("anti_dumping")
                        else "跨章归类" if original_hs[:4] != best["hs_code"][:4]
                        else ""
                    ),
                }
                optimization_logs.append(log)

                logger.info(
                    f"跨章优化: {original_name[:20]} {original_hs}({original_rate}) "
                    f"→ {best['hs_code']}({best['tax_rate']}) score={best['final_score']:.3f}"
                )
            else:
                optimized_items.append(item)

        logger.info(
            f"跨章优化完成: {len(items)} → 优化 {len(optimization_logs)} 个"
        )
        return {
            "items": optimized_items,
            "optimization_logs": optimization_logs,
        }

    def _find_best_candidate(
        self,
        original_hs: str,
        original_name: str,
        tax_data: Dict[str, Dict],
    ) -> Optional[Dict[str, Any]]:
        """在全量 tax_data 中搜索最佳候选编码"""
        original_info = tax_data.get(original_hs)
        original_rate = original_info.get("tax_rate", "") if original_info else ""

        candidates = []

        for hs_code, info in tax_data.items():
            if hs_code == original_hs:
                continue

            # 过滤条件
            if self.exclude_anti_dumping and info.get("anti_dumping"):
                continue
            if self.exclude_cert and info.get("certification_required"):
                continue

            # 计算品名相似度
            candidate_name = info.get("description_cn", "")
            similarity = self._calculate_similarity(original_name, candidate_name)

            # 章节偏好加成
            prefix4 = hs_code[:4] if len(hs_code) >= 4 else hs_code
            if prefix4 in self.preferred_chapters:
                similarity = min(1.0, similarity + self.chapter_boost)

            # 最低相似度过滤
            if similarity < self.min_similarity:
                continue

            # 计算税率得分
            tax_score = self._calculate_tax_score(original_rate, info.get("tax_rate", ""))

            # 综合得分
            final_score = tax_score * self.tax_weight + similarity * self.similarity_weight

            candidates.append({
                "hs_code": hs_code,
                "tax_rate": info.get("tax_rate", ""),
                "similarity": similarity,
                "anti_dumping": info.get("anti_dumping", False),
                "tax_score": tax_score,
                "final_score": final_score,
            })

        if not candidates:
            return None

        # 按 final_score 降序
        candidates.sort(key=lambda x: x["final_score"], reverse=True)
        best = candidates[0]

        # 判断是否比原始更好
        is_better = best["final_score"] > self.adopt_threshold

        best["is_better"] = is_better
        return best

    def _calculate_tax_score(self, original_rate: str, new_rate: str) -> float:
        """
        计算税率得分: 税率越低得分越高
        新税率 < 原税率 → 得分 > 0.5
        新税率 == 原税率 → 0.5
        新税率 > 原税率 → 得分 < 0.5
        """
        orig = self._parse_rate(original_rate)
        new = self._parse_rate(new_rate)

        if orig is None or new is None:
            return 0.5

        if new < orig:
            # 降税幅度越大得分越高
            reduction_pct = (orig - new) / max(orig, 0.01)
            return 0.5 + reduction_pct * 0.5
        elif new == orig:
            return 0.5
        else:
            return 0.0

    @staticmethod
    def _parse_rate(rate_str: str) -> Optional[float]:
        """解析税率字符串为数值"""
        if not rate_str:
            return None
        s = str(rate_str).lower().strip()
        if s in ("free", "0", "0%", ""):
            return 0.0
        try:
            return float(s.replace("%", ""))
        except (ValueError, TypeError):
            return None

    @staticmethod
    def _calculate_similarity(name1: str, name2: str) -> float:
        """
        计算文本相似度 (bigram Jaccard + 字符重叠)
        与原 optimizer.py 算法一致
        """
        if not name1 or not name2:
            return 0.0

        t1 = name1.replace(" ", "").replace("/", "").replace("-", "")
        t2 = name2.replace(" ", "").replace("/", "").replace("-", "")

        if not t1 or not t2:
            return 0.0

        # Bigram
        def bigrams(text):
            return {text[i:i+2] for i in range(len(text)-1)} if len(text) >= 2 else set(text)

        bg1, bg2 = bigrams(t1), bigrams(t2)
        bg_score = len(bg1 & bg2) / max(len(bg1 | bg2), 1) if bg1 and bg2 else 0.0

        # 单字重叠
        c1, c2 = set(t1), set(t2)
        char_score = len(c1 & c2) / max(len(c1 | c2), 1) if c1 and c2 else 0.0

        return bg_score * 0.5 + char_score * 0.5
