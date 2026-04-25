"""
HS编码优化模块
"""
import logging
from typing import List, Dict, Any, Optional

logger = logging.getLogger(__name__)


class HSCodeOptimizer:
    """HS编码优化器"""

    def __init__(self, options: Dict[str, Any] = None):
        self.options = options or {}
        self.exclude_anti_dumping = self.options.get("exclude_anti_dumping", False)
        self.min_similarity = self.options.get("min_similarity", 0.6)

    def optimize_batch(
        self,
        items: List[Dict[str, Any]],
        tax_data: Dict[str, Dict]
    ) -> Dict[str, Any]:
        """
        批量优化HS编码

        Args:
            items: 商品列表
            tax_data: 税率数据

        Returns:
            优化结果
        """
        optimized_items = []
        optimization_logs = []

        for item in items:
            result = self._optimize_single_item(item, tax_data)
            optimized_items.append(result["item"])

            if result["log"]:
                optimization_logs.append(result["log"])

        logger.info(f"优化完成，共优化 {len(optimization_logs)} 个商品")

        return {
            "items": optimized_items,
            "optimization_logs": optimization_logs
        }

    def _optimize_single_item(
        self,
        item: Dict[str, Any],
        tax_data: Dict[str, Dict]
    ) -> Dict[str, Any]:
        """优化单个商品"""
        original_hs_code = str(item.get("商品编码", ""))
        original_name = item.get("中文品名", "")

        # 获取原始税率信息
        original_tax_info = tax_data.get(original_hs_code)

        if not original_tax_info:
            # 没有税率信息，不优化
            return {"item": item, "log": None}

        # 查找候选编码（这里简化处理，实际需要更复杂的逻辑）
        candidates = self._find_candidates(original_hs_code, original_name, tax_data)

        if not candidates:
            return {"item": item, "log": None}

        # 选择最优编码
        best_candidate = self._select_best_candidate(
            original_tax_info,
            candidates,
            original_name
        )

        if best_candidate:
            # 创建优化后的商品
            optimized_item = item.copy()
            optimized_item["商品编码"] = best_candidate["hs_code"]
            optimized_item["_original_hs_code"] = original_hs_code
            optimized_item["_optimized"] = True
            optimized_item["_similarity_score"] = best_candidate["similarity"]

            # 创建日志
            log = {
                "original_name": original_name,
                "original_hs_code": original_hs_code,
                "original_tax_rate": original_tax_info.get("tax_rate", ""),
                "new_hs_code": best_candidate["hs_code"],
                "new_tax_rate": best_candidate["tax_rate"],
                "reason": "税率更低",
                "similarity_score": best_candidate["similarity"],
                "risk_warning": "⚠️ 有反倾销标记" if best_candidate.get("anti_dumping") else ""
            }

            return {"item": optimized_item, "log": log}

        return {"item": item, "log": None}

    def _find_candidates(
        self,
        original_hs_code: str,
        original_name: str,
        tax_data: Dict[str, Dict]
    ) -> List[Dict]:
        """查找候选编码"""
        candidates = []

        # 简化逻辑：查找相同前缀的编码
        prefix = original_hs_code[:4] if len(original_hs_code) >= 4 else ""

        for hs_code, info in tax_data.items():
            if hs_code == original_hs_code:
                continue

            # 前缀匹配
            if hs_code.startswith(prefix):
                # 过滤条件
                if self.exclude_anti_dumping and info.get("anti_dumping"):
                    continue

                if info.get("certification_required"):
                    continue

                # 计算相似度（简化）
                similarity = self._calculate_similarity(original_name, info.get("description_cn", ""))

                if similarity >= self.min_similarity:
                    candidates.append({
                        "hs_code": hs_code,
                        "tax_rate": info.get("tax_rate", ""),
                        "similarity": similarity,
                        "anti_dumping": info.get("anti_dumping", False)
                    })

        return candidates

    def _calculate_similarity(self, name1: str, name2: str) -> float:
        """
        计算相似度
        简化实现：基于关键词匹配
        """
        if not name1 or not name2:
            return 0.0

        # 提取关键词
        keywords1 = set(name1)
        keywords2 = set(name2)

        # 计算交集
        intersection = keywords1 & keywords2

        if not keywords1:
            return 0.0

        return len(intersection) / len(keywords1)

    def _select_best_candidate(
        self,
        original_tax_info: Dict,
        candidates: List[Dict],
        original_name: str
    ) -> Optional[Dict]:
        """选择最优候选编码"""
        if not candidates:
            return None

        # 计算综合得分
        for candidate in candidates:
            # 税率得分（Free = 1.0, 其他按比例）
            tax_score = self._calculate_tax_score(
                original_tax_info.get("tax_rate", ""),
                candidate["tax_rate"]
            )

            # 综合得分 = 税率得分(60%) + 相似度得分(40%)
            candidate["final_score"] = tax_score * 0.6 + candidate["similarity"] * 0.4

        # 按得分排序
        candidates.sort(key=lambda x: x["final_score"], reverse=True)

        # 返回得分最高的
        best = candidates[0]

        # 只有当得分明显更好时才返回
        if best["final_score"] > 0.7:
            return best

        return None

    def _calculate_tax_score(self, original_rate: str, new_rate: str) -> float:
        """计算税率得分"""
        if new_rate.lower() == "free":
            return 1.0

        # 简化处理：提取数字
        try:
            original_val = float(original_rate.replace("%", ""))
            new_val = float(new_rate.replace("%", ""))

            if new_val < original_val:
                return 1.0 - (new_val / 100)
            else:
                return 0.0
        except:
            return 0.5
