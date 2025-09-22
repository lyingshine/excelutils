"""
价格匹配器 - 处理价格匹配和更新逻辑
"""

import pandas as pd
import re
from typing import Dict, Tuple, Optional
import sys
import os

# 添加项目根目录到Python路径
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

try:
    from utils.logger import logger
    from core.data_extractor import DataExtractor
except ImportError:
    import logging
    logger = logging.getLogger(__name__)
    from data_extractor import DataExtractor


class PriceMatcher:
    """价格匹配器类，负责价格匹配和更新逻辑"""
    
    def __init__(self):
        self.data_extractor = DataExtractor()

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
                        speed_val = self.data_extractor.extract_speed_from_name(name) or ''
                        speed_val = str(speed_val).strip()

                    # 原始尺寸：优先取列值，否则从简称提取
                    size = str(row.get('尺寸', '')).strip()
                    if not size:
                        size = self.data_extractor.extract_size_from_name(name) or ''
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