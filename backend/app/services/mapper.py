"""
字段映射模块
"""
from typing import List, Dict, Any
import logging

logger = logging.getLogger(__name__)


class FieldMapper:
    """字段映射器"""

    # 标准字段映射规则
    STANDARD_FIELDS = {
        "中文品名": ["品名", "货物名称", "商品名称", "中文品名"],
        "英文品名": ["Description", "Product Name", "英文品名", "英文描述"],
        "商品编码": ["HS编码", "HS Code", "海关编码", "商品编码"],
        "材质": ["材质", "Material", "材料"],
        "用途": ["用途", "Usage", "使用用途"],
        "数量": ["数量", "Quantity", "QTY", "qty"],
        "单位": ["单位", "Unit"],
        "净重": ["净重", "Net Weight", "NW"],
        "毛重": ["毛重", "Gross Weight", "GW"],
        "箱数": ["箱数", "件数", "Cartons", "箱子数"],
    }

    def map_fields(self, parsed_data: Dict[str, Any]) -> List[Dict[str, Any]]:
        """
        映射字段到标准格式

        Args:
            parsed_data: 解析后的数据

        Returns:
            映射后的数据列表
        """
        mapped_items = []
        mapping_logs = []

        for file_data in parsed_data["files"]:
            for sheet in file_data["sheets"]:
                headers = sheet["headers"]

                # 建立字段映射关系
                field_mapping = self._build_field_mapping(headers)
                mapping_logs.extend(field_mapping["logs"])

                # 转换数据
                for row in sheet["rows"]:
                    mapped_row = self._map_row(row, field_mapping["mapping"])
                    if mapped_row:
                        mapped_items.append(mapped_row)

        logger.info(f"字段映射完成，共 {len(mapped_items)} 条数据")

        return {
            "items": mapped_items,
            "mapping_logs": mapping_logs
        }

    def _build_field_mapping(self, headers: List[str]) -> Dict[str, Any]:
        """建立字段映射关系"""
        mapping = {}
        logs = []

        for standard_field, keywords in self.STANDARD_FIELDS.items():
            for header in headers:
                if self._is_field_match(header, keywords):
                    mapping[header] = standard_field
                    logs.append({
                        "source_field": header,
                        "target_field": standard_field,
                        "confidence": 0.95,
                        "method": "keyword_match"
                    })
                    break

        return {"mapping": mapping, "logs": logs}

    def _is_field_match(self, header: str, keywords: List[str]) -> bool:
        """判断字段是否匹配"""
        header_lower = str(header).lower().strip()

        for keyword in keywords:
            keyword_lower = keyword.lower()
            if keyword_lower == header_lower or keyword_lower in header_lower:
                return True

        return False

    def _map_row(self, row: Dict[str, Any], mapping: Dict[str, str]) -> Dict[str, Any]:
        """映射单行数据"""
        mapped_row = {}

        for source_field, target_field in mapping.items():
            value = row.get(source_field)
            if value is not None and str(value).strip() != "" and str(value) != "nan":
                mapped_row[target_field] = value

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
