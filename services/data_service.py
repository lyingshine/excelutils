"""
数据服务 - 整合数据处理流程的服务层
"""

import pandas as pd
from typing import Optional
import sys
import os

# 添加项目根目录到Python路径
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

try:
    from utils.logger import logger
    from core.data_extractor import DataExtractor
    from core.data_filter import DataFilter
    from core.price_matcher import PriceMatcher
    from core.profit_calculator import ProfitCalculator
    from models.data_models import ProcessingResult
except ImportError as e:
    import logging
    logger = logging.getLogger(__name__)
    raise ImportError(f"导入模块失败: {e}")


class DataService:
    """数据服务类，整合所有数据处理流程"""
    
    def __init__(self):
        self.data_extractor = DataExtractor()
        self.data_filter = DataFilter()
        self.price_matcher = PriceMatcher()
        self.profit_calculator = ProfitCalculator()

    def import_excel_data(self, file_path: str) -> pd.DataFrame:
        """导入Excel数据"""
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
            logger.info(f"成功导入数据，共{len(df)}行，ID列已保持为字符串格式")
            return df
        except Exception as e:
            logger.error(f"导入Excel数据失败: {e}")
            raise

    def process_data(self, df: pd.DataFrame) -> pd.DataFrame:
        """处理原始数据"""
        try:
            # 复制数据避免修改原始数据
            processed_df = df.copy()

            # 删除货品ID和规格ID列
            columns_to_drop = []
            for col in processed_df.columns:
                if '货品ID' in col or '规格ID' in col:
                    columns_to_drop.append(col)

            if columns_to_drop:
                processed_df = processed_df.drop(columns=columns_to_drop)

            # 去重
            processed_df = processed_df.drop_duplicates()

            # 从简称中提取信息
            if '简称' in processed_df.columns:
                processed_df = self.data_extractor.extract_info_from_name(processed_df)

            logger.info(f"数据处理完成，共{len(processed_df)}行")
            return processed_df

        except Exception as e:
            logger.error(f"数据处理失败: {e}")
            raise

    def generate_profit_table(self, df: pd.DataFrame) -> pd.DataFrame:
        """生成毛利表"""
        try:
            # 应用数据筛选规则
            df_filtered = self.data_filter.apply_data_filtering_rules(df)
            
            # 生成毛利表
            profit_table = self.profit_calculator.generate_profit_table(df_filtered)
            
            return profit_table

        except Exception as e:
            logger.error(f"生成毛利表失败: {e}")
            raise

    def update_prices(self, original_data: pd.DataFrame, modified_profit_table: pd.DataFrame) -> pd.DataFrame:
        """根据更新后的毛利表更新价格"""
        try:
            return self.price_matcher.update_prices(original_data, modified_profit_table)
        except Exception as e:
            logger.error(f"更新价格失败: {e}")
            raise

    def get_processing_summary(self, original_count: int, processed_count: int, profit_count: int) -> ProcessingResult:
        """获取处理摘要"""
        success = profit_count > 0
        message = f"处理完成：原始数据{original_count}行 -> 处理后{processed_count}行 -> 毛利表{profit_count}行"
        
        return ProcessingResult(
            success=success,
            message=message,
            row_count=profit_count
        )