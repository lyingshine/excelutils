"""
数据筛选器 - 应用数据筛选规则
"""

import pandas as pd
from typing import List, Dict, Any
import sys
import os

# 添加项目根目录到Python路径
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

try:
    from utils.logger import logger
except ImportError:
    import logging
    logger = logging.getLogger(__name__)


class DataFilter:
    """数据筛选器类，负责应用各种数据筛选规则"""
    
    def __init__(self):
        pass

    def apply_data_filtering_rules(self, df: pd.DataFrame) -> pd.DataFrame:
        """应用数据筛选规则"""
        try:
            filtered_data = []

            # 确保必要的列存在
            required_cols = ['配置', '颜色', '尺寸', '速别', '价格', '成本']
            for col in required_cols:
                if col not in df.columns:
                    df[col] = ''

            # 为促销款产品设置默认尺寸
            if '分类' in df.columns:
                for index, row in df.iterrows():
                    if (pd.isna(row.get('尺寸')) or row.get('尺寸') == '') and '促销' in str(row.get('分类', '')):
                        df.at[index, '尺寸'] = '26寸'

            # 创建配置+颜色的组合键
            df['配置颜色组合'] = df['配置'].fillna('') + df['颜色'].fillna('')
            config_color_combos = df['配置颜色组合'].unique()

            for combo in config_color_combos:
                if pd.isna(combo) or combo == '':
                    continue

                config_data = df[df['配置颜色组合'] == combo].copy()

                # 规则2：所有配置只保留26寸的数据
                if '26寸' in config_data['尺寸'].values:
                    config_data_26 = config_data[config_data['尺寸'] == '26寸']
                    if not config_data_26.empty:
                        config_data = config_data_26
                else:
                    # 如果没有26寸，选择成本最高的尺寸
                    if '成本' in config_data.columns and not config_data['成本'].empty:
                        try:
                            config_data['成本数值'] = pd.to_numeric(config_data['成本'], errors='coerce')
                            size_max_costs = config_data.groupby('尺寸')['成本数值'].max()
                            if not size_max_costs.empty:
                                max_cost_size = size_max_costs.idxmax()
                                config_data = config_data[config_data['尺寸'] == max_cost_size]
                        except:
                            first_size = config_data['尺寸'].iloc[0] if not config_data['尺寸'].empty else ''
                            if first_size:
                                config_data = config_data[config_data['尺寸'] == first_size]

                # 规则4：选择最低价格的数据
                if '价格' in config_data.columns and not config_data.empty:
                    grouped = config_data.groupby(['配置', '尺寸', '速别'])
                    for group_key, group_data in grouped:
                        if len(group_data) > 1:
                            try:
                                group_data['价格数值'] = pd.to_numeric(group_data['价格'], errors='coerce')
                                min_price = group_data['价格数值'].min()
                                min_price_data = group_data[group_data['价格数值'] == min_price]
                                selected_row = min_price_data.iloc[0]
                            except:
                                selected_row = group_data.iloc[0]
                        else:
                            selected_row = group_data.iloc[0]
                        
                        filtered_data.append(selected_row)
                else:
                    for _, row in config_data.iterrows():
                        filtered_data.append(row)

            # 转换为DataFrame
            if filtered_data:
                result_df = pd.DataFrame(filtered_data)
                result_df = result_df.reset_index(drop=True)
                # 删除临时列
                cols_to_drop = ['配置颜色组合', '成本数值', '价格数值']
                for col in cols_to_drop:
                    if col in result_df.columns:
                        result_df = result_df.drop(columns=[col])
                return result_df
            else:
                if '配置颜色组合' in df.columns:
                    df = df.drop(columns=['配置颜色组合'])
                return df

        except Exception as e:
            logger.error(f"应用数据筛选规则时出错: {e}")
            cols_to_drop = ['配置颜色组合', '成本数值', '价格数值']
            for col in cols_to_drop:
                if col in df.columns:
                    df = df.drop(columns=[col])
            return df