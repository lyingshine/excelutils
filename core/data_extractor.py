"""
数据提取器 - 从商品名称中提取配置、尺寸、速别等信息
"""

import re
import pandas as pd
from typing import List, Tuple, Optional
import sys
import os

# 添加项目根目录到Python路径
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

try:
    from utils.logger import logger
except ImportError:
    import logging
    logger = logging.getLogger(__name__)


class DataExtractor:
    """数据提取器类，负责从商品名称中提取各种信息"""
    
    def __init__(self):
        self.size_patterns = [r'24寸', r'26寸', r'27\.5寸']
        self.speed_pattern = r'(\d+速)'
        self.color_patterns = [r'黑', r'白', r'红', r'蓝', r'绿', r'黄', r'灰', r'银', r'金', r'粉', r'紫', r'橙']

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