"""
文件生成模块
"""
import openpyxl
from openpyxl.styles import Font
from datetime import datetime
from typing import List, Dict, Any
from pathlib import Path
import logging

logger = logging.getLogger(__name__)


class FileGenerator:
    """文件生成器"""

    def generate(
        self,
        optimized_data: List[Dict[str, Any]],
        template_path: str,
        output_dir: str
    ) -> str:
        """
        生成清关文件

        Args:
            optimized_data: 优化后的数据
            template_path: 模板文件路径
            output_dir: 输出目录

        Returns:
            生成的文件路径
        """
        try:
            logger.info(f"开始生成清关文件，共 {len(optimized_data)} 条数据")

            # 加载模板
            wb = openpyxl.load_workbook(template_path)
            ws = wb.active

            # 找到数据起始行
            start_row = self._find_data_start_row(ws)

            # 填充数据
            for idx, item in enumerate(optimized_data):
                row_num = start_row + idx
                self._fill_row(ws, row_num, item)

            # 保存文件
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            output_filename = f"清关_{timestamp}.xlsx"
            output_path = Path(output_dir) / output_filename

            wb.save(output_path)

            logger.info(f"清关文件生成成功: {output_path}")

            return str(output_path)

        except Exception as e:
            logger.error(f"生成清关文件失败: {str(e)}")
            raise

    def _find_data_start_row(self, ws) -> int:
        """找到数据起始行（表头下一行）"""
        # 简单策略：找到第一行包含"品名"的行，然后返回下一行
        for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=20), start=1):
            for cell in row:
                if cell.value and "品名" in str(cell.value):
                    return row_idx + 1

        # 默认从第2行开始
        return 2

    def _fill_row(self, ws, row_num: int, item: Dict[str, Any]):
        """填充一行数据"""
        # 字段到列的映射（需要根据实际模板调整）
        field_column_mapping = {
            "中文品名": 1,
            "英文品名": 2,
            "商品编码": 3,
            "材质": 4,
            "用途": 5,
            "数量": 6,
            "单位": 7,
            "净重": 8,
            "毛重": 9,
            "箱数": 10,
        }

        for field, col_idx in field_column_mapping.items():
            value = item.get(field)

            if value is not None:
                ws.cell(row_num, col_idx).value = value
            else:
                # 缺失字段标记
                ws.cell(row_num, col_idx).value = "待补充"
                ws.cell(row_num, col_idx).font = Font(color="FF0000")

    def generate_log_file(
        self,
        logs: Dict[str, Any],
        output_dir: str
    ) -> str:
        """
        生成日志文件

        Args:
            logs: 日志数据
            output_dir: 输出目录

        Returns:
            日志文件路径
        """
        try:
            logger.info("开始生成日志文件")

            wb = openpyxl.Workbook()

            # Sheet1: 替换日志
            ws1 = wb.active
            ws1.title = "替换日志"
            ws1.append(["原品名", "原HS编码", "原税率", "新HS编码", "新税率", "替换原因", "风险提示"])

            for log in logs.get("optimization_logs", []):
                ws1.append([
                    log.get("original_name", ""),
                    log.get("original_hs_code", ""),
                    log.get("original_tax_rate", ""),
                    log.get("new_hs_code", ""),
                    log.get("new_tax_rate", ""),
                    log.get("reason", ""),
                    log.get("risk_warning", "")
                ])

            # Sheet2: 数据冲突日志
            ws2 = wb.create_sheet("数据冲突日志")
            ws2.append(["字段", "文档A的值", "文档B的值", "处理方式"])

            for conflict in logs.get("conflicts", []):
                ws2.append([
                    conflict.get("field", ""),
                    conflict.get("value_a", ""),
                    conflict.get("value_b", ""),
                    "已空着，请手动填写"
                ])

            # Sheet3: 字段映射日志
            ws3 = wb.create_sheet("字段映射日志")
            ws3.append(["输入文档字段", "映射到清关字段", "置信度"])

            for mapping in logs.get("mapping_logs", []):
                confidence = mapping.get("confidence", 0)
                ws3.append([
                    mapping.get("source_field", ""),
                    mapping.get("target_field", ""),
                    f"{confidence*100:.0f}%"
                ])

            # 保存文件
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            output_filename = f"日志_{timestamp}.xlsx"
            output_path = Path(output_dir) / output_filename

            wb.save(output_path)

            logger.info(f"日志文件生成成功: {output_path}")

            return str(output_path)

        except Exception as e:
            logger.error(f"生成日志文件失败: {str(e)}")
            raise
