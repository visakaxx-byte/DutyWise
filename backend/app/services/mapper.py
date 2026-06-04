"""
字段映射模块
"""
from typing import List, Dict, Any
import logging
import re

logger = logging.getLogger(__name__)

# 多源回退的最低匹配分数阈值（低于此分数的备用源不纳入）
FALLBACK_MIN_SCORE = 0.55


class FieldMapper:
    """字段映射器"""

    # 标准字段映射规则
    STANDARD_FIELDS = {
        "中文品名": ["品名", "货物名称", "商品名称", "中文品名"],
        "英文品名": ["Description", "Product Name", "英文品名", "英文描述"],
        "商品编码": ["HS编码", "HS Code", "海关编码", "商品编码"],
        "材质": ["中英文材质", "中文材质", "材质", "Material", "材料"],
        "用途": ["中英文用途", "用途", "Usage", "使用用途"],
        "数量": ["总数量", "数量", "Quantity", "QTY", "qty"],
        "单位": ["单位", "Unit"],
        "箱数": ["箱数量", "箱数", "件数", "Cartons", "箱子数"],
        "总价": ["申报价值", "总价", "总金额", "申报总价", "invoice总价", "Total Value", "Subtotal", "Invoice Value", "申报金额", "清关申报金额"],
        "币制": ["申报价值币别", "币制", "币别", "Currency"],
        "单价": ["申报单价", "单价", "Unit Price", "Unit Value"],
        "净重": ["净重", "Net Weight", "NW", "实重"],
        "毛重": ["毛重", "Gross Weight", "GW", "总抛重"],
        "原产国": ["原产国", "Country of Origin", "生产国", "原产地"],
    }

    def map_fields(self, parsed_data: Dict[str, Any]) -> List[Dict[str, Any]]:
        """
        映射字段到标准格式
        """
        mapped_items = []
        mapping_logs = []

        for file_data in parsed_data["files"]:
            for sheet in file_data["sheets"]:
                headers = sheet["headers"]

                # 跳过明确标记为无产品字段的sheet
                if sheet.get("has_product_fields") is False:
                    logger.debug(f"跳过无产品字段的sheet: {sheet.get('sheet_name', 'unknown')}")
                    continue

                # 建立字段映射关系
                field_mapping = self._build_field_mapping(headers)
                mapping_logs.extend(field_mapping["logs"])

                # 找出运单号列的原始列名（用于后续分组填充）
                waybill_source = None
                for h in headers:
                    if str(h).strip() in ("运单号码", "提单号码", "柜号"):
                        waybill_source = h
                        break

                # 第一遍：映射所有行，同时收集分组信息
                rows_with_data = []  # (mapped_row, source_row)
                for row in sheet["rows"]:
                    mapped_row = self._map_row(row, field_mapping["mapping"])
                    if mapped_row:
                        # 暂存运单号用于分组
                        if waybill_source and waybill_source in row:
                            wb_val = row[waybill_source]
                            if wb_val is not None and str(wb_val).strip() not in ("", "nan"):
                                mapped_row["_waybill"] = str(wb_val).strip()
                        rows_with_data.append(mapped_row)

                # 第二遍：按运单号分组填充缺失的箱数/总价/币制
                self._fill_by_waybill(rows_with_data)

                mapped_items.extend(rows_with_data)

        # 全局默认值填充
        for item in mapped_items:
            if "原产国" not in item:
                item["原产国"] = "CN"
            if "单位" not in item:
                item["单位"] = "PCS"

        logger.info(f"字段映射完成，共 {len(mapped_items)} 条数据")

        return {
            "items": mapped_items,
            "mapping_logs": mapping_logs
        }

    def _fill_by_waybill(self, items: List[Dict[str, Any]]):
        """按运单号分组，将组内首行的箱数/总价/币制填充到缺失行"""
        # 收集每个运单号的首个有效值
        group_values: Dict[str, Dict[str, Any]] = {}
        for item in items:
            wb = item.get("_waybill")
            if not wb:
                continue
            if wb not in group_values:
                group_values[wb] = {}
            for field in ("箱数", "总价", "币制"):
                if field not in group_values[wb] and field in item:
                    group_values[wb][field] = item[field]

        # 对缺失行进行填充
        filled_count = 0
        for item in items:
            wb = item.get("_waybill")
            if not wb or wb not in group_values:
                # 无运单号的行：数量=1时 箱数=1
                if "箱数" not in item and "数量" in item:
                    try:
                        qty = float(item["数量"])
                        if qty == 1:
                            item["箱数"] = 1
                            filled_count += 1
                    except (ValueError, TypeError):
                        pass
                if "币制" not in item and "总价" in item:
                    item["币制"] = "USD"
                    filled_count += 1
                continue

            gv = group_values[wb]
            for field in ("箱数", "总价", "币制"):
                if field not in item and field in gv:
                    item[field] = gv[field]
                    filled_count += 1

        # 清理临时字段
        for item in items:
            item.pop("_waybill", None)

        logger.info(f"分组填充完成，共填充 {filled_count} 个缺失字段")

    def _build_field_mapping(self, headers: List[str]) -> Dict[str, Any]:
        """建立字段映射关系（支持多源回退：首选列为空时尝试备用列）"""
        mapping: Dict[str, List[str]] = {}  # {standard_field: [source1, source2, ...]}
        logs = []
        used_primaries: set = set()  # 已被其他标准字段占用的首选列

        for standard_field, keywords in self.STANDARD_FIELDS.items():
            # 收集所有匹配的源列及得分（同时记录匹配的关键词位置，用于同分时破序）
            scored = []
            for header in headers:
                best_score = 0.0
                best_kw_idx = 999
                for kw_idx, kw in enumerate(keywords):
                    s = self._match_single(header, kw)
                    if s > best_score or (s == best_score and kw_idx < best_kw_idx):
                        best_score = s
                        best_kw_idx = kw_idx
                if best_score > 0:
                    scored.append((header, best_score, best_kw_idx))
            # 按得分降序，同分时按关键词列表先后顺序（靠前的优先）
            scored.sort(key=lambda x: (-x[1], x[2]))

            if not scored:
                continue

            # 过滤：排除语义不兼容的源列（如 "英文品名" 不能作为 "中文品名"）
            filtered = self._filter_incompatible(standard_field, scored)

            if not filtered:
                continue

            # 收集未占用首选的列作为候选源（仅保留 score >= FALLBACK_MIN_SCORE）
            sources = []
            for header, score in filtered:
                if header in used_primaries:
                    continue
                if not sources:
                    # 首选列必须达到最低分数
                    if score >= FALLBACK_MIN_SCORE:
                        sources.append(header)
                elif score >= FALLBACK_MIN_SCORE:
                    # 备用列也需达到最低分数
                    sources.append(header)

            if sources:
                # 标记首选列为已占用
                used_primaries.add(sources[0])
                mapping[standard_field] = sources
                # 日志记录首选匹配
                logs.append({
                    "source_field": sources[0],
                    "target_field": standard_field,
                    "confidence": min(filtered[0][1], 0.95),
                    "method": "keyword_match"
                })
                if len(sources) > 1:
                    logger.debug(f"字段 {standard_field} 多源映射: {sources}")

        return {"mapping": mapping, "logs": logs}

    def _filter_incompatible(self, standard_field: str, scored: List[tuple]) -> List[tuple]:
        """过滤语义不兼容的候选源列（如中文/英文语言冲突）
        输入: [(header, score, kw_idx), ...] 3元组
        输出: [(header, score), ...] 2元组（丢弃kw_idx）"""
        header_items = [(h, s) for h, s, _ in scored]
        # 中文类标准字段：排除含"英文/English"的源列
        if "中文" in standard_field:
            header_items = [(h, s) for h, s in header_items
                           if "英文" not in str(h) and "English" not in str(h)]
        # 英文类标准字段：排除含"中文/Chinese"的源列
        if "英文" in standard_field:
            header_items = [(h, s) for h, s in header_items
                           if "中文" not in str(h) and "Chinese" not in str(h)]
        # 总价：排除含"币别/币制/Currency"的源列（避免把币制当成总价）
        if standard_field == "总价":
            header_items = [(h, s) for h, s in header_items
                           if not re.search(r'币别|币制|Currency', str(h))]
        return header_items

    def _match_score(self, header: str, keywords: List[str]) -> float:
        """返回匹配得分：精确匹配 > 边界匹配 > 包含匹配
        中文文本禁用前缀/后缀匹配（避免 "申报价值币别" 误匹配 "总价"）"""
        best = 0.0
        for kw in keywords:
            s = self._match_single(header, kw)
            if s > best:
                best = s
        return best

    def _match_single(self, header: str, keyword: str) -> float:
        """对单个关键词计算匹配得分"""
        header_lower = str(header).lower().strip()
        header_str = str(header).strip()
        header_clean = header_str.replace("\n", "").replace("\r", "")
        kw_lower = keyword.lower()
        kw_str = str(keyword)

        if header_lower == kw_lower:
            return 1.0

        # 清理后完全匹配（移除换行后）
        if header_clean.lower() == kw_lower:
            return 1.0

        # 检测是否含中文
        is_cjk = bool(re.search(r'[\u4e00-\u9fff]', kw_str)) or bool(re.search(r'[\u4e00-\u9fff]', header_str))

        if is_cjk:
            # 中文文本：前缀/后缀优先（高置信），再尝试包含匹配
            # 前/后缀匹配仅当剩余部分是明显非中文内容（英文/数字/括号）时生效
            if header_lower.startswith(kw_lower):
                remainder = header_lower[len(kw_lower):]
                if not remainder or re.match(r'^[\s\n\r0-9a-z()（）/.\-]+$', remainder, re.IGNORECASE):
                    return 0.85
            if header_lower.endswith(kw_lower):
                prefix = header_lower[:len(header_lower)-len(kw_lower)]
                if re.match(r'^[\s\n\r0-9a-z()（）/.\-]+$', prefix, re.IGNORECASE):
                    return 0.85
            if kw_lower in header_lower:
                ratio = len(kw_lower) / max(len(header_lower), 1)
                return 0.4 + ratio * 0.3
        else:
            if header_lower.startswith(kw_lower) or header_lower.endswith(kw_lower):
                return 0.9
            elif kw_lower in header_lower:
                ratio = len(kw_lower) / max(len(header_lower), 1)
                return 0.4 + ratio * 0.3

        return 0.0

    def _map_row(self, row: Dict[str, Any], mapping: Dict[str, List[str]]) -> Dict[str, Any]:
        """映射单行数据（支持多源回退：首选列为空时尝试备用列）"""
        mapped_row = {}

        for standard_field, source_fields in mapping.items():
            for source_field in source_fields:
                value = row.get(source_field)
                if value is not None and str(value).strip() != "" and str(value) != "nan":
                    mapped_row[standard_field] = value
                    break

        # 只返回有中文品名的行
        if "中文品名" in mapped_row:
            return mapped_row

        return None

    def resolve_conflicts(self, items: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        解决数据冲突
        简单策略：记录冲突，留空让用户手动填写
        """
        conflicts = []

        # 这里可以实现更复杂的冲突检测逻辑
        # 例如：检查同一商品在不同文件中的数据是否一致

        return {
            "items": items,
            "conflicts": conflicts
        }
