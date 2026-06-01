"""
低申报策略模块

对输入的申报价值进行调整:
  - 数量 = 原始数量 × quantity_multiplier
  - 单价 = 原始单价 × price_multiplier
  - 总价 = 数量 × 单价
  - 总价 > total_cap → 单价 = total_cap / 数量 (封顶)
  - 保护字段: 毛重/箱数 永不修改
"""

import logging
from typing import Any, Dict, List, Optional

logger = logging.getLogger(__name__)


class DevaluationStrategy:
    """低申报策略: 乘数法调整申报价值 + 总额封顶"""

    def __init__(self, config: Dict[str, Any] = None):
        cfg = config or {}
        self.price_multiplier = cfg.get("price_multiplier", 0.15)
        self.quantity_multiplier = cfg.get("quantity_multiplier", 1.0)
        self.total_cap = cfg.get("total_cap", 15000)
        self.protected_fields = set(cfg.get("protected_fields", ["毛重", "箱数"]))

    def apply(self, items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        对商品列表应用低申报策略。

        操作:
          1. 数量 = 原始值 × quantity_multiplier
          2. 单价 = 原始值 × price_multiplier
          3. 总价 = 数量 × 单价
          4. 总价封顶限制
          5. 保护 毛重、箱数 不被修改
        """
        if not items:
            return items

        total_declared = 0.0

        for item in items:
            # 保存原始值用于日志
            item["_original_单位"] = item.get("单价")
            item["_original_数量"] = item.get("数量")
            item["_original_总价"] = item.get("总价")

            # Step 1: 数量调整
            orig_qty = self._to_float(item.get("数量"))
            if orig_qty is not None and self.quantity_multiplier != 1.0:
                new_qty = round(orig_qty * self.quantity_multiplier)
                item["数量"] = max(1, new_qty)

            # Step 2: 单价调整
            orig_price = self._to_float(item.get("单价"))
            if orig_price is not None:
                new_price = round(orig_price * self.price_multiplier, 2)
                item["单价"] = max(0.01, new_price)

            # Step 3: 总价 = 数量 × 单价
            qty = self._to_float(item.get("数量"))
            price = self._to_float(item.get("单价"))
            if qty is not None and price is not None:
                item["总价"] = round(qty * price, 2)

            total_declared += self._to_float(item.get("总价")) or 0

        # Step 4: 总价封顶 (按比例缩减所有商品的单价以达成封顶)
        if total_declared > self.total_cap:
            ratio = self.total_cap / total_declared
            logger.info(f"总价 ${total_declared:,.2f} 超过封顶 ${self.total_cap:,}，按比例 {ratio:.3f} 缩减")
            for item in items:
                price = self._to_float(item.get("单价"))
                if price is not None:
                    item["单价"] = round(price * ratio, 2)
                qty = self._to_float(item.get("数量"))
                p = self._to_float(item.get("单价"))
                if qty is not None and p is not None:
                    item["总价"] = round(qty * p, 2)
            # 重新计算
            total_declared = sum(self._to_float(i.get("总价")) or 0 for i in items)

        logger.info(f"低申报调整完成，申报总价: ${total_declared:,.2f}")
        return items

    @staticmethod
    def _to_float(value: Any) -> Optional[float]:
        """安全转换为 float"""
        if value is None:
            return None
        try:
            return float(value)
        except (ValueError, TypeError):
            return None
