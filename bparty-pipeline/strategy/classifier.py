"""
品名归类模块 - 细名→宽泛类别 + HS编码赋值 + 同类合并 + 提单品名保护

category_map 格式 (config.yaml):
  category_map:
    "关键词1":
      category: "宽泛类别名"
      hs: "3926400090"
    "关键词2":
      category: "宽泛类别名"
      hs: "3926400090"

提单保护: 提单上列出的品名会在归类+合并时被保留为独立行，
  不会被重命名为类别名，也不会与其他行合并。
"""

import logging
from typing import Dict, List, Any, Optional, Set

logger = logging.getLogger(__name__)

# B/L 品名关键词匹配时的英语停用词
BL_STOP_WORDS = {
    "the", "and", "for", "with", "set", "new", "used",
    "of", "or", "in", "to", "a", "an", "is", "at", "on",
    "no", "not", "per", "as", "by", "be",
}


class ProductClassifier:
    """品名归类器: 关键词重命名 + HS赋值 + 同类型合并"""

    def __init__(self, config: Dict[str, Any] = None):
        cfg = config or {}
        self.mode = cfg.get("mode", "keyword")
        self.category_map_raw: Dict[str, Any] = cfg.get("category_map", {})
        self.merge_config = cfg.get("merge", {})
        self.merge_enabled = self.merge_config.get("enabled", True)
        self.group_by = self.merge_config.get("group_by", ["中文品名", "商品编码", "材质"])
        self.sum_fields = self.merge_config.get("sum_fields", ["数量", "净重", "毛重", "箱数", "总价"])

        # 提单品名列表 (外部传入)
        self.bl_products: List[str] = []

        # 规范化 category_map: 统一转为 {category, hs} 格式
        self.category_map: Dict[str, Dict[str, str]] = {}
        for key, val in self.category_map_raw.items():
            if isinstance(val, str):
                self.category_map[key] = {"category": val, "hs": ""}
            elif isinstance(val, dict):
                self.category_map[key] = {
                    "category": val.get("category", ""),
                    "hs": val.get("hs", ""),
                }
            else:
                logger.warning(f"category_map 条目格式错误: {key!r}")

        # 按关键词长度降序排序，优先匹配更长的关键词
        self._sorted_keys = sorted(self.category_map.keys(), key=len, reverse=True)

    def set_bl_products(self, products: List[str]) -> None:
        """设置提单品名列表（在 classify 之前调用）"""
        self.bl_products = [p.strip() for p in products if p.strip()]

    def classify(self, items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        对商品列表执行品名归类。

        Args:
            items: 映射后的商品列表 (含 中文品名、商品编码 等)

        Returns:
            归类并合并后的商品列表
        """
        if not items:
            return items

        # Step 0: 标记提单保护项 + 无匹配品名创建兜底条目
        if self.bl_products:
            self._mark_bl_protected(items)

        # Step 1: 关键词重命名 + HS编码赋值（跳过受保护项）
        matched_count = 0
        unmatched_names: Set[str] = set()
        for item in items:
            # 提单保护项: 不改名、不换HS
            if item.get("_bl_protected"):
                continue

            name = item.get("中文品名", "")
            if name and self._sorted_keys:
                matched = self._match_keyword(str(name))
                if matched:
                    item["_original_name"] = item.get("中文品名", "")
                    item["中文品名"] = matched["category"]
                    if matched.get("hs"):
                        item["_original_hs_code"] = item.get("商品编码", "")
                        item["商品编码"] = matched["hs"]
                    matched_count += 1
                else:
                    unmatched_names.add(str(name))

        if unmatched_names:
            logger.warning(
                f"归类: {matched_count}/{len(items)} 条已匹配, "
                f"{len(unmatched_names)} 个品名未匹配: {list(unmatched_names)[:10]}"
            )

        # Step 2: 同类型合并（受保护项独立成组，不与其他行合并）
        if self.merge_enabled:
            items = self._merge_similar(items)

        return items

    # ── 提单保护 ──

    def _mark_bl_protected(self, items: List[Dict[str, Any]]) -> None:
        """
        遍历 items，标记与提单品名匹配的行为 _bl_protected。

        匹配策略（仅对英文品名，避免中文名误匹配）:
          Tier 1: 完整子串匹配（提单品名是 item 英文名的子串，或反之）
          Tier 2: 去空格后完全匹配
          Tier 3: 关键词覆盖率 >= 50% 且 >= 2 个关键词命中

        未匹配的提单品名会被创建为合成条目，追加到 items 中。
        """
        matched_bl_products: Set[str] = set()

        for bl_product in self.bl_products:
            keywords = self._extract_bl_keywords(bl_product)
            if not keywords:
                continue

            bl_lower = bl_product.lower()
            bl_nospace = bl_lower.replace(" ", "")
            matched = False

            for item in items:
                if item.get("_bl_protected"):
                    continue

                en_name = str(item.get("英文品名", "")).lower().strip()
                if not en_name:
                    continue

                # Tier 1: 完整子串匹配
                if bl_lower in en_name or en_name in bl_lower:
                    self._protect_item(item, bl_product)
                    matched_bl_products.add(bl_product)
                    matched = True
                    break

                # Tier 2: 去空格后完全匹配
                en_nospace = en_name.replace(" ", "")
                if bl_nospace == en_nospace:
                    self._protect_item(item, bl_product)
                    matched_bl_products.add(bl_product)
                    matched = True
                    break

                # Tier 3: 关键词覆盖率 >= 50% 且 >= 2 个关键词命中
                en_words = set(en_name.replace("-", " ").replace("/", " ").split())
                overlap = sum(1 for kw in keywords if kw in en_words)
                coverage = overlap / len(keywords)
                if overlap >= 2 and coverage >= 0.5:
                    self._protect_item(item, bl_product)
                    matched_bl_products.add(bl_product)
                    matched = True
                    break

            if not matched:
                logger.warning(
                    f"提单品名 '{bl_product}' 在清单中未找到匹配项，"
                    f"将创建兜底条目（关键词: {keywords}）"
                )
                synthetic = self._create_synthetic_item(bl_product)
                logger.info(
                    f"提单合成条目: {bl_product} → "
                    f"材质={synthetic.get('材质', '-')}, "
                    f"HS={synthetic.get('商品编码', '-')}"
                )
                items.append(synthetic)

        if matched_bl_products:
            logger.info(
                f"提单保护: {len(matched_bl_products)}/{len(self.bl_products)} 个品名已匹配 "
                f"→ {sorted(matched_bl_products)}"
            )

    def _protect_item(self, item: Dict[str, Any], bl_product: str) -> None:
        """标记 item 为 B/L 保护项"""
        item["_bl_protected"] = True
        item["_bl_source"] = bl_product
        logger.info(
            f"提单保护: 英文名 '{item.get('英文品名', '')[:50]}' ↔ {bl_product}"
        )

    def _create_synthetic_item(self, bl_product_name: str) -> Dict[str, Any]:
        """为未匹配的提单品名创建合成条目（带材质推断 + HS编码尝试匹配）"""
        material = self._infer_material(bl_product_name)

        # 从品名中剥离材质词，得到纯品名
        clean_name = self._strip_material_from_name(bl_product_name, material)

        # 尝试从 category_map 匹配 HS 编码
        hs_code = ""
        category = clean_name
        clean_lower = clean_name.lower()
        for key in self._sorted_keys:
            if key.lower() in clean_lower:
                entry = self.category_map[key]
                if entry.get("hs"):
                    hs_code = entry["hs"]
                    category = entry["category"]
                    break

        return {
            "中文品名": category if hs_code else clean_name,
            "英文品名": clean_name,
            "材质": material,
            "商品编码": hs_code,
            "用途": "HOME",
            "数量": 1,
            "净重": 0,
            "毛重": 0,
            "箱数": 1,
            "单位": "PCS",
            "币制": "USD",
            "总价": 0,
            "原产国": "CN",
            "_bl_protected": True,
            "_bl_source": bl_product_name,
            "_synthetic": True,
        }

    @staticmethod
    def _strip_material_from_name(name: str, material: str) -> str:
        """从品名中剥离材质词，返回纯品名"""
        if not material:
            return name.strip()
        # 不区分大小写替换材质词
        words = name.split()
        mat_lower = material.lower()
        filtered = [w for w in words if w.lower() != mat_lower]
        return " ".join(filtered).strip() if filtered else name.strip()

    @staticmethod
    def _infer_material(name: str) -> str:
        """从品名推断材质"""
        name_lower = name.lower()
        material_keywords = [
            ("plastic", "Plastic"),
            ("silicone", "Silicone"),
            ("metal", "Metal"),
            ("steel", "Steel"),
            ("stainless", "Stainless Steel"),
            ("aluminum", "Aluminum"),
            ("wood", "Wood"),
            ("wooden", "Wood"),
            ("glass", "Glass"),
            ("ceramic", "Ceramic"),
            ("rubber", "Rubber"),
            ("leather", "Leather"),
            ("fabric", "Fabric"),
            ("cotton", "Cotton"),
            ("paper", "Paper"),
            ("bamboo", "Bamboo"),
            ("copper", "Copper"),
            ("iron", "Iron"),
            ("brass", "Brass"),
        ]
        for kw, mat in material_keywords:
            if kw in name_lower:
                return mat
        return ""

    def _extract_bl_keywords(self, bl_name: str) -> List[str]:
        """
        从提单品名中提取可用于匹配的关键词。

        去掉停用词，保留 >=3 个字符的英文单词，统一小写。
        例如 "PLASTIC MOBILE PHONE STAND" → ["plastic", "mobile", "phone", "stand"]
        """
        words = bl_name.lower().replace("-", " ").replace("/", " ").split()
        keywords = [w for w in words if len(w) >= 3 and w not in BL_STOP_WORDS]
        return keywords

    # ── 关键词匹配 ──

    def _match_keyword(self, name: str) -> Optional[Dict[str, str]]:
        """在品名中查找关键词，返回 {category, hs} 或 None"""
        for key in self._sorted_keys:
            if key in name:
                return self.category_map[key]
        return None

    # ── 合并逻辑 ──

    def _merge_similar(self, items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        合并相同 (品名, HS编码, 材质) 的行。

        提单保护项独立成组，永不与其他行合并。
        数值字段累加，非数值/非sum字段取首行值。
        """
        if not items:
            return items

        groups: Dict[tuple, List[Dict[str, Any]]] = {}
        group_order: List[tuple] = []

        for item in items:
            is_protected = item.get("_bl_protected", False)

            if is_protected:
                # 受保护项: 使用唯一键（包含 id() 或序号）
                # 简单做法: 用 item 的内存地址作为唯一键
                group_key = ("__BL__", id(item))
            else:
                # 构建分组键
                key_parts = []
                for field in self.group_by:
                    val = item.get(field, "")
                    key_parts.append(str(val).strip() if val is not None else "")
                group_key = tuple(key_parts)

            if group_key not in groups:
                groups[group_key] = []
                group_order.append(group_key)
            groups[group_key].append(item)

        merged = []
        for gk in group_order:
            group = groups[gk]
            if len(group) == 1:
                merged.append(group[0])
            else:
                merged_item = self._merge_group(group)
                merged.append(merged_item)
                logger.debug(f"合并 {len(group)} 行: {gk[0] if gk[0] != '__BL__' else '提单保护项'}")

        logger.info(f"合并: {len(items)} → {len(merged)} 条")
        return merged

    def _merge_group(self, group: List[Dict[str, Any]]) -> Dict[str, Any]:
        """合并一组相同商品，数值字段求和"""
        result = group[0].copy()

        merged_count = len(group)
        if merged_count > 1:
            result["_merged_from"] = merged_count

        for field in self.sum_fields:
            total = 0.0
            has_value = False
            for item in group:
                val = item.get(field)
                if val is not None:
                    try:
                        total += float(val)
                        has_value = True
                    except (ValueError, TypeError):
                        pass
            if has_value:
                result[field] = total

        return result
