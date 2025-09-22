"""
毛利计算器 - 处理毛利表生成和计算逻辑
"""

import pandas as pd
from typing import List, Dict, Any
import sys
import os

# 添加项目根目录到Python路径
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

try:
    from utils.logger import logger
    from models.data_models import ProfitTableRow
except ImportError:
    import logging
    logger = logging.getLogger(__name__)


class ProfitCalculator:
    """毛利计算器类，负责毛利表生成和相关计算"""
    
    def __init__(self):
        pass

    def generate_profit_table(self, df: pd.DataFrame) -> pd.DataFrame:
        """生成毛利表"""
        try:
            # 创建毛利表数据
            profit_data = []
            
            # 首先处理缺少尺寸信息的产品，显示在最上方
            if '简称' in df.columns:
                no_size_mask = ~df['简称'].str.contains('寸', na=False)
                no_size_data = df[no_size_mask]
                if not no_size_data.empty:
                    for _, row in no_size_data.iterrows():
                        config_name = str(row.get('简称', '缺少尺寸信息')) if pd.notna(row.get('简称')) else '缺少尺寸信息'
                        self._add_profit_row(profit_data, row, config_name)

            # 处理不适用筛选逻辑的数据（配置为空的数据）
            if '配置' in df.columns:
                no_config_mask = df['配置'].isna() | (df['配置'] == '')
                no_config_data = df[no_config_mask]
                # 排除已经处理过的缺少尺寸信息的产品
                if '简称' in no_config_data.columns:
                    has_size_mask = no_config_data['简称'].str.contains('寸', na=False)
                    no_config_data = no_config_data[has_size_mask]
                if not no_config_data.empty:
                    for _, row in no_config_data.iterrows():
                        config_name = str(row.get('简称', '未知配置')) if pd.notna(row.get('简称')) else '未知配置'
                        self._add_profit_row(profit_data, row, config_name)

            # 处理有配置的数据 - 配置名称直接用配置+颜色拼接
            if '配置' in df.columns:
                has_config_mask = df['配置'].notna() & (df['配置'] != '')
                has_config_data = df[has_config_mask]
                # 排除已经处理过的缺少尺寸信息的产品
                if '简称' in has_config_data.columns:
                    has_size_mask = has_config_data['简称'].str.contains('寸', na=False)
                    has_config_data = has_config_data[has_size_mask]
                if not has_config_data.empty:
                    # 创建配置+颜色的组合键
                    has_config_data = has_config_data.copy()
                    has_config_data['配置颜色组合'] = has_config_data['配置'].fillna('') + has_config_data['颜色'].fillna('')
                    config_color_combos = has_config_data['配置颜色组合'].unique()

                    for combo in config_color_combos:
                        if pd.isna(combo) or combo == '':
                            continue

                        combo_data = has_config_data[has_config_data['配置颜色组合'] == combo]
                        if '速别' in combo_data.columns:
                            speeds = combo_data['速别'].unique()
                        else:
                            speeds = ['未知速别']

                        for speed in speeds:
                            if pd.isna(speed) or speed == '':
                                continue

                            speed_data = combo_data[combo_data['速别'] == speed]
                            if not speed_data.empty:
                                selected_row = speed_data.iloc[0]
                                # 配置名称直接用配置+颜色拼接
                                config_with_color = str(combo)
                                self._add_profit_row(profit_data, selected_row, config_with_color)

            # 转换为DataFrame
            profit_df = pd.DataFrame(profit_data)
            if not profit_df.empty:
                # 确保缺少尺寸信息的数据显示在最上方
                if '配置' in profit_df.columns:
                    no_size_rows = profit_df[profit_df['配置'].str.contains('缺少尺寸信息', na=False)]
                    no_config_rows = profit_df[profit_df['配置'].str.contains('未知配置', na=False)]
                    has_config_rows = profit_df[~profit_df['配置'].str.contains('未知配置|缺少尺寸信息', na=False)]
                    
                    if not has_config_rows.empty:
                        has_config_rows = self._sort_profit_table(has_config_rows)
                    
                    profit_df = pd.concat([no_size_rows, no_config_rows, has_config_rows], ignore_index=True)

            logger.info(f"毛利表生成完成，共{len(profit_df)}行")
            return profit_df

        except Exception as e:
            logger.error(f"生成毛利表失败: {e}")
            raise

    def _add_profit_row(self, profit_data: List[Dict], row: pd.Series, config: str):
        """添加毛利表行数据"""
        price = row.get('价格', '')
        cost = row.get('成本', '')
        profit = row.get('毛利', '')
        profit_rate = row.get('毛利率', '')
        name = row.get('简称', '')
        color = row.get('颜色', '')

        # 格式化价格
        price = self._format_price(price)
        cost = self._format_price(cost)
        profit = self._format_profit(profit, price, cost)
        profit_rate = self._format_profit_rate(profit_rate, profit, price)

        # 配置列包含颜色信息 - 特殊处理渐变色
        config_with_color = config
        if color and str(color).strip() != '':
            # 对于渐变色产品，保持配置名称不变，颜色信息单独处理
            if '渐变' in str(color):
                config_with_color = config
            else:
                config_with_color = f"{config}{color}"
        
        profit_data.append({
            '配置': config_with_color,
            '速别': row.get('速别', ''),
            '简称': name if pd.notna(name) else '',
            '价格': price,
            '成本': cost,
            '快递': '30.00',
            '毛利润': profit,
            '毛利率': profit_rate
        })

    def _format_price(self, value) -> str:
        """格式化价格"""
        if pd.notna(value) and value != '':
            try:
                return f"{float(value):.2f}"
            except:
                return str(value)
        return '0.00'

    def _format_profit(self, profit, price, cost) -> str:
        """格式化毛利润"""
        if pd.isna(profit) or profit == '':
            try:
                profit_val = float(price) - float(cost) if price != '' and cost != '' else 0
                return f"{profit_val:.2f}"
            except:
                return '0.00'
        return profit

    def _format_profit_rate(self, profit_rate, profit, price) -> str:
        """格式化毛利率"""
        if pd.isna(profit_rate) or profit_rate == '':
            try:
                if float(price) > 0:
                    profit_rate_val = (float(profit) / float(price)) * 100
                    return f"{profit_rate_val:.2f}%"
                else:
                    return '0.00%'
            except:
                return '0.00%'
        else:
            try:
                rate_val = float(profit_rate)
                if rate_val > 1:
                    return f"{rate_val:.2f}%"
                else:
                    return f"{rate_val * 100:.2f}%"
            except:
                return str(profit_rate)

    def _sort_profit_table(self, df: pd.DataFrame) -> pd.DataFrame:
        """对毛利表进行排序"""
        config_max_prices = {}
        for config in df['配置'].unique():
            config_data = df[df['配置'] == config]
            max_price = 0
            for _, row in config_data.iterrows():
                try:
                    price_val = float(str(row['价格']).replace(',', '')) if row['价格'] and row['价格'] != '' else 0
                    max_price = max(max_price, price_val)
                except:
                    pass
            config_max_prices[config] = max_price

        sorted_configs = sorted(config_max_prices.items(), key=lambda x: x[1])
        sorted_data = []
        for config, _ in sorted_configs:
            config_data = df[df['配置'] == config]
            config_data = config_data.sort_values('速别')
            sorted_data.append(config_data)

        return pd.concat(sorted_data, ignore_index=True) if sorted_data else df