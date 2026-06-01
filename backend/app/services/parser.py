"""
文档解析模块
"""
import pandas as pd
from typing import List, Dict, Any
from pathlib import Path
import logging
import re

logger = logging.getLogger(__name__)

# 产品报关相关字段关键词（用于判断文件是否包含产品申报信息）
PRODUCT_FIELD_KEYWORDS = [
    "品名", "HS", "编码", "材质", "用途", "单位", "净重", "毛重",
    "申报", "海关", "数量", "单价", "总价", "币制", "Description",
    "Material", "Quantity", "Product", "商品"
]


class DocumentParser:
    """文档解析器"""

    def parse_files(self, file_paths: List[str]) -> Dict[str, Any]:
        """
        解析多个Excel文件

        Args:
            file_paths: 文件路径列表

        Returns:
            解析结果
        """
        results = []

        for file_path in file_paths:
            try:
                logger.info(f"开始解析文件: {file_path}")
                file_data = self._parse_single_file(file_path)
                results.append(file_data)
            except Exception as e:
                logger.error(f"解析文件失败 {file_path}: {str(e)}")
                raise

        return {"files": results}

    def _parse_single_file(self, file_path: str) -> Dict[str, Any]:
        """解析单个文件"""
        filename = Path(file_path).name
        sheets_data = []

        # 读取所有sheet
        excel_file = pd.ExcelFile(file_path)

        for sheet_name in excel_file.sheet_names:
            # 跳过明显是说明/注释的sheet
            if any(kw in str(sheet_name) for kw in ["说明", "注释", "Note", "说明页"]):
                logger.info(f"跳过说明sheet: {sheet_name}")
                continue

            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)

            # 检测表头行
            header_row = self._detect_header_row(df)

            if header_row is not None:
                # 重新读取，指定表头行
                df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row)

                # 过滤掉完全空的行
                df = df.dropna(how='all')

                # 转换为字典列表
                rows = df.to_dict('records')

                # 检查是否包含产品报关字段
                if rows and not self._has_product_fields(list(df.columns)):
                    logger.warning(
                        f"文件 {filename} 的sheet \"{sheet_name}\" "
                        f"不包含产品报关字段（品名/HS编码/申报价值等），"
                        f"可能是物流清单而非报关明细。"
                        f"已识别的列: {list(df.columns)[:10]}..."
                    )
                    # 仍然保留解析结果（可能部分字段有用），但在结果中标记
                    sheets_data.append({
                        "sheet_name": sheet_name,
                        "headers": list(df.columns),
                        "rows": rows,
                        "row_count": len(rows),
                        "has_product_fields": False,
                    })
                else:
                    sheets_data.append({
                        "sheet_name": sheet_name,
                        "headers": list(df.columns),
                        "rows": rows,
                        "row_count": len(rows),
                        "has_product_fields": True,
                    })
            else:
                logger.warning(f"文件 {filename} 的sheet \"{sheet_name}\" 未找到有效表头")
                # 尝试作为无表头数据处理
                df_all = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
                df_all = df_all.dropna(how='all')
                if len(df_all) > 0:
                    cols = list(df_all.iloc[0])
                    sheets_data.append({
                        "sheet_name": sheet_name,
                        "headers": [str(c) for c in cols],
                        "rows": [],
                        "row_count": 0,
                        "has_product_fields": False,
                    })

        return {
            "filename": filename,
            "sheets": sheets_data
        }

    def _detect_header_row(self, df: pd.DataFrame) -> int:
        """
        检测表头行
        策略：
        1. 找到非空列数最多的行（但至少 >= 5 列）
        2. 这样做会自动跳过顶部仅含2-3列的meta信息行
        """
        best_row = None
        best_count = 0

        # 限制扫描前50行，避免在超大表格中浪费时间
        max_scan = min(len(df), 50)

        for idx in range(max_scan):
            row = df.iloc[idx]
            non_null_count = row.notna().sum()

            # 排除关键元数据行：这些行只有2-3列（提单号/柜号等meta信息）
            # 真正的表头行通常有5列以上的列名
            if non_null_count >= 5 and non_null_count > best_count:
                # 额外检查：该行是否包含典型的表头关键词
                row_values = [str(v).strip() for v in row if pd.notna(v)]
                has_header_keywords = any(
                    kw in val for kw in PRODUCT_FIELD_KEYWORDS
                    for val in row_values
                )
                # 如果找到产品关键词行，直接返回（这是最佳表头）
                if has_header_keywords:
                    logger.debug(f"在行 {idx} 找到产品字段表头: {row_values[:5]}...")
                    return idx

                # 记录备选（非空列最多但不是产品字段的行）
                if non_null_count > best_count:
                    best_count = non_null_count
                    best_row = idx

        # 如果没找到产品关键词行，回退到非空列最多的行
        if best_row is not None:
            logger.debug(f"回退表头行 {best_row} (非空列数: {best_count})")
            return best_row

        # 最后的回退：第一个 >= 3 列的行
        for idx in range(max_scan):
            if df.iloc[idx].notna().sum() >= 3:
                return idx

        return None

    def _has_product_fields(self, headers: List[str]) -> bool:
        """检查表头是否包含产品报关相关字段"""
        header_str = " ".join(str(h) for h in headers if h is not None)
        return any(
            kw in header_str
            for kw in PRODUCT_FIELD_KEYWORDS
        )
