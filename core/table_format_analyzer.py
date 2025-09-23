"""
表格格式分析器 - 分析数据特征并决定毛利表格式
"""

import pandas as pd
from typing import Dict, List, Tuple
import sys
import os

# 添加项目根目录到Python路径
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

try:
    from utils.logger import logger
except ImportError:
    import logging
    logger = logging.getLogger(__name__)


class TableFormatAnalyzer:
    """表格格式分析器类，负责分析数据特征并决定毛利表的显示格式"""
    
    def __init__(self):
        pass

    def analyze_data_characteristics(self, df: pd.DataFrame) -> Dict:
        """分析数据特征，返回分析结果"""
        try:
            result = {
                'size_count': 0,
                'speed_count': 0,
                'unique_sizes': [],
                'unique_speeds': [],
                'should_use_size_format': False,
                'format_type': 'speed_based'  # 默认基于速别的格式
            }
            
            # 统计尺寸项数量
            if '尺寸' in df.columns:
                unique_sizes = df['尺寸'].dropna().unique()
                # 过滤掉空字符串
                unique_sizes = [size for size in unique_sizes if size and str(size).strip() != '']
                result['unique_sizes'] = list(unique_sizes)
                result['size_count'] = len(unique_sizes)
            
            # 统计速别项数量
            if '速别' in df.columns:
                unique_speeds = df['速别'].dropna().unique()
                # 过滤掉空字符串
                unique_speeds = [speed for speed in unique_speeds if speed and str(speed).strip() != '']
                result['unique_speeds'] = list(unique_speeds)
                result['speed_count'] = len(unique_speeds)
            
            # 判断是否应该使用尺寸格式
            if result['size_count'] > result['speed_count']:
                result['should_use_size_format'] = True
                result['format_type'] = 'size_based'
                logger.info(f"检测到尺寸项({result['size_count']})多于速别项({result['speed_count']})，将使用尺寸格式")
            else:
                logger.info(f"使用默认速别格式：尺寸项({result['size_count']})，速别项({result['speed_count']})")
            
            return result
            
        except Exception as e:
            logger.error(f"分析数据特征失败: {e}")
            return {
                'size_count': 0,
                'speed_count': 0,
                'unique_sizes': [],
                'unique_speeds': [],
                'should_use_size_format': False,
                'format_type': 'speed_based'
            }

    def should_use_size_based_format(self, df: pd.DataFrame) -> bool:
        """判断是否应该使用基于尺寸的格式"""
        analysis = self.analyze_data_characteristics(df)
        return analysis['should_use_size_format']

    def get_format_columns(self, use_size_format: bool) -> List[str]:
        """根据格式类型返回列顺序"""
        if use_size_format:
            # 尺寸格式：配置列包含速别信息，尺寸列替换速别列
            return ['配置', '尺寸', '简称', '价格', '成本', '快递', '毛利润', '毛利率']
        else:
            # 默认速别格式
            return ['配置', '速别', '简称', '价格', '成本', '快递', '毛利润', '毛利率']

    def format_config_name_with_speed(self, config: str, speed: str) -> str:
        """将速别信息拼接到配置名中"""
        if not config or pd.isna(config):
            config = ''
        if not speed or pd.isna(speed) or str(speed).strip() == '':
            return str(config)
        
        config_str = str(config)
        speed_str = str(speed)
        
        # 如果配置名已经包含速别信息，不重复添加
        if speed_str in config_str:
            return config_str
        
        # 拼接速别到配置名
        return f"{config_str}{speed_str}"