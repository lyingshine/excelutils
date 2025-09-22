"""
应用程序主类 - 整合所有功能的应用层
"""

import pandas as pd
from typing import Optional
import sys
import os

# 添加项目根目录到Python路径
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

try:
    from utils.logger import logger
    from services.data_service import DataService
    from services.excel_service import ExcelService
    from models.data_models import ProcessingResult
except ImportError as e:
    import logging
    logger = logging.getLogger(__name__)
    raise ImportError(f"导入模块失败: {e}")


class Application:
    """应用程序主类，提供完整的业务流程"""
    
    def __init__(self):
        self.data_service = DataService()
        self.excel_service = ExcelService()
        
        # 数据存储
        self.original_data: Optional[pd.DataFrame] = None
        self.processed_data: Optional[pd.DataFrame] = None
        self.profit_table_data: Optional[pd.DataFrame] = None
        self.updated_data: Optional[pd.DataFrame] = None

    def import_data(self, file_path: str) -> ProcessingResult:
        """导入Excel数据"""
        try:
            self.original_data = self.data_service.import_excel_data(file_path)
            
            return ProcessingResult(
                success=True,
                message=f"数据导入成功，共{len(self.original_data)}行数据",
                data=self.original_data,
                row_count=len(self.original_data)
            )
        except Exception as e:
            logger.error(f"导入数据失败: {e}")
            return ProcessingResult(
                success=False,
                message=f"导入数据失败: {str(e)}"
            )

    def process_and_generate_profit_table(self) -> ProcessingResult:
        """处理数据并生成毛利表"""
        if self.original_data is None:
            return ProcessingResult(
                success=False,
                message="请先导入数据"
            )
        
        try:
            # 处理数据
            self.processed_data = self.data_service.process_data(self.original_data)
            
            # 生成毛利表
            self.profit_table_data = self.data_service.generate_profit_table(self.processed_data)
            
            return ProcessingResult(
                success=True,
                message=f"毛利表生成完成，共{len(self.profit_table_data)}行",
                data=self.profit_table_data,
                row_count=len(self.profit_table_data)
            )
        except Exception as e:
            logger.error(f"处理数据失败: {e}")
            return ProcessingResult(
                success=False,
                message=f"处理数据失败: {str(e)}"
            )

    def export_profit_table(self, file_path: str) -> ProcessingResult:
        """导出毛利表"""
        if self.profit_table_data is None:
            return ProcessingResult(
                success=False,
                message="请先生成毛利表"
            )
        
        return self.excel_service.export_profit_table(
            file_path, self.profit_table_data, self.original_data
        )

    def import_modified_profit_table(self, file_path: str) -> ProcessingResult:
        """导入修改后的毛利表并更新价格"""
        if self.original_data is None:
            return ProcessingResult(
                success=False,
                message="请先导入原始数据"
            )
        
        try:
            # 读取修改后的毛利表
            result = self.excel_service.import_excel_file(file_path)
            if not result.success:
                return result
            
            modified_profit_table = result.data
            
            # 更新价格
            self.updated_data = self.data_service.update_prices(
                self.original_data.copy(), modified_profit_table
            )
            
            # 更新毛利表数据
            self.profit_table_data = modified_profit_table
            
            return ProcessingResult(
                success=True,
                message="已成功导入更新后的毛利表并更新价格",
                data=self.updated_data,
                row_count=len(self.updated_data)
            )
        except Exception as e:
            logger.error(f"导入修改后毛利表失败: {e}")
            return ProcessingResult(
                success=False,
                message=f"导入修改后毛利表失败: {str(e)}"
            )

    def export_updated_data(self, file_path: str) -> ProcessingResult:
        """导出更新后的原始数据"""
        export_data = self.updated_data if self.updated_data is not None else self.original_data
        
        if export_data is None:
            return ProcessingResult(
                success=False,
                message="没有可导出的数据"
            )
        
        return self.excel_service.export_original_data(file_path, export_data)

    def get_data_summary(self) -> dict:
        """获取数据摘要"""
        return {
            'original_count': len(self.original_data) if self.original_data is not None else 0,
            'processed_count': len(self.processed_data) if self.processed_data is not None else 0,
            'profit_table_count': len(self.profit_table_data) if self.profit_table_data is not None else 0,
            'updated_count': len(self.updated_data) if self.updated_data is not None else 0,
            'has_original_data': self.original_data is not None,
            'has_processed_data': self.processed_data is not None,
            'has_profit_table': self.profit_table_data is not None,
            'has_updated_data': self.updated_data is not None
        }