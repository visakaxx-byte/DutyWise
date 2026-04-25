"""
文档解析模块
"""
import pandas as pd
from typing import List, Dict, Any
from pathlib import Path
import logging

logger = logging.getLogger(__name__)


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
            df = pd.read_excel(file_path, sheet_name=sheet_name)

            # 检测表头行
            header_row = self._detect_header_row(df)

            if header_row is not None:
                # 重新读取，指定表头行
                df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row)

                # 转换为字典列表
                rows = df.to_dict('records')

                sheets_data.append({
                    "sheet_name": sheet_name,
                    "headers": list(df.columns),
                    "rows": rows,
                    "row_count": len(rows)
                })

        return {
            "filename": filename,
            "sheets": sheets_data
        }

    def _detect_header_row(self, df: pd.DataFrame) -> int:
        """
        检测表头行
        简单策略：找到第一个包含多个非空值的行
        """
        for idx, row in df.iterrows():
            non_null_count = row.notna().sum()
            if non_null_count >= 3:  # 至少3个非空列
                return idx
        return 0
