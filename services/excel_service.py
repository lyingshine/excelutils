"""
Excel服务 - 整合Excel导入导出功能的服务层
"""

import pandas as pd
from typing import Optional
import sys
import os

# 添加项目根目录到Python路径
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

try:
    from utils.logger import logger
    from exporters.excel_exporter import ExcelExporter
    from models.data_models import ProcessingResult
except ImportError as e:
    import logging
    logger = logging.getLogger(__name__)
    raise ImportError(f"导入模块失败: {e}")


class ExcelService:
    """Excel服务类，整合Excel相关操作"""
    
    def __init__(self):
        self.excel_exporter = ExcelExporter()

    def export_profit_table(self, file_path: str, profit_table: pd.DataFrame, 
                          original_data: Optional[pd.DataFrame] = None) -> ProcessingResult:
        """导出毛利表"""
        try:
            self.excel_exporter.export_profit_table(file_path, profit_table, original_data)
            
            return ProcessingResult(
                success=True,
                message=f"毛利表已成功导出到: {file_path}",
                row_count=len(profit_table)
            )
        except Exception as e:
            logger.error(f"导出毛利表失败: {e}")
            return ProcessingResult(
                success=False,
                message=f"导出毛利表失败: {str(e)}"
            )

    def export_original_data(self, file_path: str, data: pd.DataFrame) -> ProcessingResult:
        """导出原始数据"""
        try:
            self.excel_exporter.export_original_data(file_path, data)
            
            return ProcessingResult(
                success=True,
                message=f"原始数据已成功导出到: {file_path}",
                row_count=len(data)
            )
        except Exception as e:
            logger.error(f"导出原始数据失败: {e}")
            return ProcessingResult(
                success=False,
                message=f"导出原始数据失败: {str(e)}"
            )

    def import_excel_file(self, file_path: str) -> ProcessingResult:
        """导入Excel文件"""
        try:
            # 指定ID列为字符串类型，防止精度丢失
            dtype_dict = {}
            
            # 预读取列名来识别ID列
            temp_df = pd.read_excel(file_path, nrows=0)
            for col in temp_df.columns:
                if 'ID' in col or 'id' in col or '货品ID' in col or '规格ID' in col:
                    dtype_dict[col] = str
            
            # 重新读取数据，指定ID列为字符串类型
            df = pd.read_excel(file_path, dtype=dtype_dict)
            
            return ProcessingResult(
                success=True,
                message=f"成功导入数据，共{len(df)}行，ID列已保持为字符串格式",
                data=df,
                row_count=len(df)
            )
        except Exception as e:
            logger.error(f"导入Excel数据失败: {e}")
            return ProcessingResult(
                success=False,
                message=f"导入Excel数据失败: {str(e)}"
            )