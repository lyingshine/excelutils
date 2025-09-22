"""
数据处理器 - 重构版本
使用服务层架构，保持API兼容性
"""

import pandas as pd
import sys
import os

# 添加项目根目录到Python路径
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.append(project_root)

from utils.logger import logger
from services.data_service import DataService

class DataProcessor:
    """数据处理类，负责数据导入、处理和毛利表生成
    
    重构版本：使用服务层架构，保持API完全兼容
    """
    
    def __init__(self):
        """初始化数据处理器"""
        # 使用服务层
        self.data_service = DataService()
        
        # 保持原有属性以确保兼容性
        self.size_patterns = [r'24寸', r'26寸', r'27\.5寸']
        self.speed_pattern = r'(\d+速)'
        self.color_patterns = [r'黑', r'白', r'红', r'蓝', r'绿', r'黄', r'灰', r'银', r'金', r'粉', r'紫', r'橙']

    def import_excel_data(self, file_path: str) -> pd.DataFrame:
        """导入Excel数据"""
        return self.data_service.import_excel_data(file_path)

    def process_data(self, df: pd.DataFrame) -> pd.DataFrame:
        """处理原始数据"""
        return self.data_service.process_data(df)

    def generate_profit_table(self, df: pd.DataFrame) -> pd.DataFrame:
        """生成毛利表"""
        return self.data_service.generate_profit_table(df)

    def update_prices(self, original_data: pd.DataFrame, modified_profit_table: pd.DataFrame) -> pd.DataFrame:
        """根据更新后的毛利表更新价格"""
        return self.data_service.update_prices(original_data, modified_profit_table)

    # ==================== 向后兼容方法 ====================
    
    def extract_info_from_name(self, df: pd.DataFrame) -> pd.DataFrame:
        """从简称中提取分类、配置、尺寸、速别信息"""
        return self.data_service.data_extractor.extract_info_from_name(df)

    def apply_data_filtering_rules(self, df: pd.DataFrame) -> pd.DataFrame:
        """应用数据筛选规则"""
        return self.data_service.data_filter.apply_data_filtering_rules(df)

    def extract_config_from_name(self, name: str, category: str) -> str:
        """从商品简称中提取配置"""
        return self.data_service.data_extractor.extract_config_from_name(name, category)

    def extract_speed_from_name(self, name: str) -> str:
        """从商品简称中提取速别"""
        return self.data_service.data_extractor.extract_speed_from_name(name)

    def extract_size_from_name(self, name: str) -> str:
        """从商品简称中提取尺寸"""
        return self.data_service.data_extractor.extract_size_from_name(name)

    def remove_size_from_name(self, name: str) -> str:
        """从商品简称中移除尺寸信息"""
        return self.data_service.data_extractor.remove_size_from_name(name)

    # ==================== 委托给服务层的方法 ====================
    
    def _add_profit_row(self, profit_data: list, row: pd.Series, config: str):
        """添加毛利表行数据"""
        return self.data_service.profit_calculator._add_profit_row(profit_data, row, config)

    def _format_price(self, value) -> str:
        """格式化价格"""
        return self.data_service.profit_calculator._format_price(value)

    def _format_profit(self, profit, price, cost) -> str:
        """格式化毛利润"""
        return self.data_service.profit_calculator._format_profit(profit, price, cost)

    def _format_profit_rate(self, profit_rate, profit, price) -> str:
        """格式化毛利率"""
        return self.data_service.profit_calculator._format_profit_rate(profit_rate, profit, price)

    def _sort_profit_table(self, df: pd.DataFrame) -> pd.DataFrame:
        """对毛利表进行排序"""
        return self.data_service.profit_calculator._sort_profit_table(df)

    def _create_price_mapping(self, profit_table: pd.DataFrame) -> dict:
        """创建价格映射"""
        return self.data_service.price_matcher._create_price_mapping(profit_table)

    def _create_price_mapping_with_index(self, profit_table: pd.DataFrame) -> dict:
        """创建价格映射（包含行号）"""
        return self.data_service.price_matcher._create_price_mapping_with_index(profit_table)

    def _calculate_new_profit_rate(self, data: pd.DataFrame) -> pd.DataFrame:
        """计算基于新价格的毛利率"""
        return self.data_service.price_matcher._calculate_new_profit_rate(data)