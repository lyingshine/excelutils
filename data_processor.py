
import pandas as pd
import re
from typing import Dict, List, Tuple, Optional
import sys
import os

# 添加项目根目录到Python路径
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

try:
    from utils.logger import logger
except ImportError:
    import logging
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
    logger = logging.getLogger(__name__)

class DataProcessor:
    """数据处理类，负责数据导入、处理和毛利表生成"""
    
    def __init__(self):
        self.size_patterns = [r'24寸', r'26寸', r'27\.5寸']
        self.speed_pattern = r'(\d+速)'
        self.color_patterns = [r'黑', r'白', r'红', r'蓝', r'绿', r'黄', r'灰', r'银', r'金', r'粉', r'紫', r'橙']

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
                processed_df = self.extract_info_from_name(processed_df)

            logger.info(f"数据处理完成，共{len(processed_df)}行")
            return processed_df

        except Exception as e:
            logger.error(f"数据处理失败: {e}")
            raise

    def extract_info_from_name(self, df: pd.DataFrame) -> pd.DataFrame:
        """从简称中提取分类、配置、尺寸、速别信息"""
        # 新增列
        df['提取的分类'] = ''
        df['配置'] = ''
        df['尺寸'] = ''
        df['速别'] = ''
        df['颜色'] = ''

        for index, row in df.iterrows():
            name = str(row.get('简称', ''))
            category = str(row.get('分类', ''))

            if not name or name == 'nan':
                continue

            # 提取分类
            df.at[index, '提取的分类'] = category

            # 查找尺寸
            size_match = None
            size_start = -1
            size_end = -1

            for pattern in self.size_patterns:
                match = re.search(pattern, name)
                if match:
                    size_match = match.group()
                    size_start = match.start()
                    size_end = match.end()
                    break

            # 查找速别
            speed_match = re.search(self.speed_pattern, name)
            speed_text = ''
            speed_start = -1
            speed_end = -1
            if speed_match:
                speed_text = speed_match.group()
                speed_start = speed_match.start()
                speed_end = speed_match.end()
                df.at[index, '速别'] = speed_text

            if size_match:
                df.at[index, '尺寸'] = size_match
                
                # 构建配置名：除了分类、尺寸、速别外的所有文字
                config_parts = []
                
                # 分类前的部分
                if category in name:
                    category_start = name.find(category)
                    if category_start > 0:
                        before_category = name[:category_start].strip()
                        if before_category:
                            config_parts.append(before_category)
                    
                    # 分类后到尺寸前的部分
                    category_end = category_start + len(category)
                    between_part = name[category_end:size_start].strip()
                    if between_part:
                        config_parts.append(between_part)
                else:
                    # 没有分类，取尺寸前的部分
                    before_size = name[:size_start].strip()
                    if before_size:
                        config_parts.append(before_size)
                
                # 尺寸后的部分（排除速别）
                after_size = name[size_end:].strip()
                if after_size:
                    if speed_match and speed_start >= size_end:
                        # 尺寸后到速别前的部分
                        between_size_speed = name[size_end:speed_start].strip()
                        if between_size_speed:
                            config_parts.append(between_size_speed)
                        
                        # 速别后的部分
                        after_speed = name[speed_end:].strip()
                        if after_speed:
                            config_parts.append(after_speed)
                    else:
                        # 没有速别或速别在尺寸前，直接添加尺寸后的部分
                        config_parts.append(after_size)
                
                # 拼接配置名
                df.at[index, '配置'] = ''.join(config_parts)

            else:
                # 如果没有尺寸信息，整个简称（除了分类和速别）都是配置
                config_parts = []
                
                if category in name:
                    # 分类前的部分
                    category_start = name.find(category)
                    if category_start > 0:
                        before_category = name[:category_start].strip()
                        if before_category:
                            config_parts.append(before_category)
                    
                    # 分类后的部分（排除速别）
                    category_end = category_start + len(category)
                    after_category = name[category_end:].strip()
                    
                    if speed_match:
                        if speed_start > category_end:
                            # 分类后到速别前的部分
                            between_part = name[category_end:speed_start].strip()
                            if between_part:
                                config_parts.append(between_part)
                        
                        # 速别后的部分
                        after_speed = name[speed_end:].strip()
                        if after_speed:
                            config_parts.append(after_speed)
                    else:
                        # 没有速别，直接添加分类后的部分
                        if after_category:
                            config_parts.append(after_category)
                else:
                    # 没有分类，整个简称（除了速别）都是配置
                    if speed_match:
                        # 速别前的部分
                        before_speed = name[:speed_start].strip()
                        if before_speed:
                            config_parts.append(before_speed)
                        
                        # 速别后的部分
                        after_speed = name[speed_end:].strip()
                        if after_speed:
                            config_parts.append(after_speed)
                    else:
                        # 没有速别，整个简称都是配置
                        config_parts.append(name.strip())
                
                # 拼接配置名
                df.at[index, '配置'] = ''.join(config_parts)
        
        return df



    def generate_profit_table(self, df: pd.DataFrame) -> pd.DataFrame:
        """生成毛利表"""
        try:
            # 应用数据筛选规则
            df_filtered = self.apply_data_filtering_rules(df)

            # 创建毛利表数据
            profit_data = []
            
            # 首先处理缺少尺寸信息的产品，显示在最上方
            if '简称' in df_filtered.columns:
                no_size_mask = ~df_filtered['简称'].str.contains('寸', na=False)
                no_size_data = df_filtered[no_size_mask]
                if not no_size_data.empty:
                    for _, row in no_size_data.iterrows():
                        config_name = str(row.get('简称', '缺少尺寸信息')) if pd.notna(row.get('简称')) else '缺少尺寸信息'
                        self._add_profit_row(profit_data, row, config_name)

            # 处理不适用筛选逻辑的数据（配置为空的数据）
            if '配置' in df_filtered.columns:
                no_config_mask = df_filtered['配置'].isna() | (df_filtered['配置'] == '')
                no_config_data = df_filtered[no_config_mask]
                # 排除已经处理过的缺少尺寸信息的产品
                if '简称' in no_config_data.columns:
                    has_size_mask = no_config_data['简称'].str.contains('寸', na=False)
                    no_config_data = no_config_data[has_size_mask]
                if not no_config_data.empty:
                    for _, row in no_config_data.iterrows():
                        config_name = str(row.get('简称', '未知配置')) if pd.notna(row.get('简称')) else '未知配置'
                        self._add_profit_row(profit_data, row, config_name)

            # 处理有配置的数据 - 配置名称直接用配置+颜色拼接
            if '配置' in df_filtered.columns:
                has_config_mask = df_filtered['配置'].notna() & (df_filtered['配置'] != '')
                has_config_data = df_filtered[has_config_mask]
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

    def update_prices(self, original_data: pd.DataFrame, modified_profit_table: pd.DataFrame) -> pd.DataFrame:
        """根据更新后的毛利表更新价格"""
        try:
            updated_data = original_data.copy()

            # 初始化修改后价格与未匹配标记
            updated_data['修改后价格'] = None
            updated_data['未匹配'] = False

            # 创建价格映射（key: 简称|速别 -> (价格, 毛利表行号)）
            price_mapping = self._create_price_mapping_with_index(modified_profit_table)

            # 统计匹配情况
            matched_count = 0
            unmatched_count = 0

            logger.info("开始价格匹配，详细日志如下：")
            logger.info("=" * 80)

            # 更新价格逻辑：
            # 1. 直接用简称匹配
            # 2. 匹配失败的，把尺寸改为26寸再进行匹配
            # 3. 最后把27.5寸的再加20元
            for index, row in updated_data.iterrows():
                try:
                    name = str(row.get('简称', '')).strip()
                    if not name or name == 'nan':
                        updated_data.at[index, '未匹配'] = True
                        unmatched_count += 1
                        logger.warning(f"原始数据第{index+1}行：简称为空，跳过匹配")
                        continue

                    # 速别：优先取列值，否则从简称提取
                    speed_val = str(row.get('速别', '')).strip()
                    if not speed_val:
                        speed_val = self.extract_speed_from_name(name) or ''
                        speed_val = str(speed_val).strip()

                    # 原始尺寸：优先取列值，否则从简称提取
                    size = str(row.get('尺寸', '')).strip()
                    if not size:
                        size = self.extract_size_from_name(name) or ''
                        size = str(size).strip()

                    base_price = None
                    matched_key = None
                    profit_table_row = None
                    match_method = None
                    
                    # 第一步：直接用简称匹配
                    key = f"{name}|{speed_val}" if speed_val else name
                    if key in price_mapping:
                        base_price, profit_table_row = price_mapping[key]
                        matched_key = key
                        match_method = "直接匹配"
                    
                    # 第二步：匹配失败的，把尺寸改为26寸再进行匹配
                    if base_price is None:
                        name_26 = re.sub(r'24寸|27\.5寸', '26寸', name)
                        key_26 = f"{name_26}|{speed_val}" if speed_val else name_26
                        if key_26 in price_mapping:
                            base_price, profit_table_row = price_mapping[key_26]
                            matched_key = key_26
                            match_method = "尺寸规范化匹配"
                        else:
                            # 尝试不带速别的键（兜底）
                            if name_26 in price_mapping:
                                base_price, profit_table_row = price_mapping[name_26]
                                matched_key = name_26
                                match_method = "尺寸规范化匹配(无速别)"

                    if base_price is not None:
                        # 第三步：最后把27.5寸的再加20元
                        final_price = float(base_price)
                        if size == '27.5寸':
                            final_price += 20.0
                            size_adjustment = " (+20元 for 27.5寸)"
                        else:
                            size_adjustment = ""
                        
                        updated_data.at[index, '修改后价格'] = final_price
                        updated_data.at[index, '未匹配'] = False
                        matched_count += 1
                        
                        # 输出匹配成功日志
                        logger.info(f"✓ 原始数据第{index+1}行 [{name}|{speed_val}|{size}] "
                                  f"-> 毛利表第{profit_table_row+1}行 [{matched_key}] "
                                  f"匹配方式: {match_method}, 价格: {base_price} -> {final_price}{size_adjustment}")
                    else:
                        # 未匹配
                        updated_data.at[index, '修改后价格'] = None
                        updated_data.at[index, '未匹配'] = True
                        unmatched_count += 1
                        
                        # 输出未匹配日志
                        logger.warning(f"✗ 原始数据第{index+1}行 [{name}|{speed_val}|{size}] 未找到匹配项")

                except Exception as e:
                    logger.error(f"处理原始数据第{index+1}行时出错: {e}")
                    updated_data.at[index, '修改后价格'] = None
                    updated_data.at[index, '未匹配'] = True
                    unmatched_count += 1
                    continue

            logger.info("=" * 80)
            logger.info(f"价格更新完成！匹配成功: {matched_count}行, 未匹配: {unmatched_count}行")
            
            # 计算新的毛利率
            updated_data = self._calculate_new_profit_rate(updated_data)
            
            return updated_data

        except Exception as e:
            logger.error(f"更新价格时出错: {e}")
            raise

    def _create_price_mapping(self, profit_table: pd.DataFrame) -> Dict[str, float]:
        """创建价格映射：key = 简称|速别 -> 26寸价格（毛利表默认26寸）"""
        price_mapping: Dict[str, float] = {}
        for _, row in profit_table.iterrows():
            name = str(row.get('简称', '')).strip()
            speed = str(row.get('速别', '')).strip()
            price = row.get('价格', '')
            if name and pd.notna(price) and str(price).strip():
                try:
                    price_val = float(str(price).replace(',', ''))
                    key = f"{name}|{speed}" if speed else name
                    price_mapping[key] = price_val
                except:
                    continue
        return price_mapping

    def _create_price_mapping_with_index(self, profit_table: pd.DataFrame) -> Dict[str, Tuple[float, int]]:
        """创建价格映射（包含行号）：key = 简称|速别 -> (价格, 毛利表行号)"""
        price_mapping: Dict[str, Tuple[float, int]] = {}
        for index, row in profit_table.iterrows():
            name = str(row.get('简称', '')).strip()
            speed = str(row.get('速别', '')).strip()
            price = row.get('价格', '')
            if name and pd.notna(price) and str(price).strip():
                try:
                    price_val = float(str(price).replace(',', ''))
                    key = f"{name}|{speed}" if speed else name
                    price_mapping[key] = (price_val, index)
                except:
                    continue
        return price_mapping

    def extract_config_from_name(self, name: str, category: str) -> str:
        """从商品简称中提取配置"""
        for pattern in self.size_patterns:
            match = re.search(pattern, name)
            if match and category in name:
                category_end = name.find(category) + len(category)
                size_start = match.start()
                config = name[category_end:size_start]
                return config
        return ''

    def extract_speed_from_name(self, name: str) -> str:
        """从商品简称中提取速别"""
        speed_match = re.search(self.speed_pattern, name)
        return speed_match.group() if speed_match else ''

    def extract_size_from_name(self, name: str) -> str:
        """从商品简称中提取尺寸"""
        for pattern in self.size_patterns:
            match = re.search(pattern, name)
            if match:
                return match.group()
        return ''

    def remove_size_from_name(self, name: str) -> str:
        """从商品简称中移除尺寸信息"""
        # 不再移除尺寸信息，保持原始名称
        return name

    def _calculate_new_profit_rate(self, data: pd.DataFrame) -> pd.DataFrame:
        """计算基于新价格的毛利率"""
        try:
            # 添加新毛利率列
            data['新毛利率'] = None
            
            # 快递费固定为30元
            express_fee = 30.0
            
            for index, row in data.iterrows():
                try:
                    new_price = row.get('修改后价格')
                    cost = row.get('成本')
                    
                    # 如果有新价格且有成本，计算新毛利率
                    if pd.notna(new_price) and pd.notna(cost) and new_price is not None:
                        try:
                            new_price_val = float(new_price)
                            cost_val = float(cost)
                            
                            if new_price_val > 0:
                                # 毛利率 = (售价 - 成本 - 快递费) / 售价 * 100%
                                profit_rate = ((new_price_val - cost_val - express_fee) / new_price_val) * 100
                                data.at[index, '新毛利率'] = f"{profit_rate:.2f}%"
                            else:
                                data.at[index, '新毛利率'] = "0.00%"
                        except (ValueError, TypeError):
                            data.at[index, '新毛利率'] = "计算错误"
                    else:
                        # 没有新价格时，使用原价格计算（如果存在）
                        original_price = row.get('价格')
                        if pd.notna(original_price) and pd.notna(cost):
                            try:
                                original_price_val = float(original_price)
                                cost_val = float(cost)
                                
                                if original_price_val > 0:
                                    # 毛利率 = (售价 - 成本 - 快递费) / 售价 * 100%
                                    profit_rate = ((original_price_val - cost_val - express_fee) / original_price_val) * 100
                                    data.at[index, '新毛利率'] = f"{profit_rate:.2f}%"
                                else:
                                    data.at[index, '新毛利率'] = "0.00%"
                            except (ValueError, TypeError):
                                data.at[index, '新毛利率'] = "计算错误"
                        else:
                            data.at[index, '新毛利率'] = "无法计算"
                            
                except Exception as e:
                    logger.warning(f"计算第{index+1}行毛利率时出错: {e}")
                    data.at[index, '新毛利率'] = "计算错误"
                    continue
            
            logger.info("新毛利率计算完成（已包含快递费30元）")
            return data
            
        except Exception as e:
            logger.error(f"计算新毛利率时出错: {e}")
            return data