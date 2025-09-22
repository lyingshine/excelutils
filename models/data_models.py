"""
数据模型定义
"""

from dataclasses import dataclass
from typing import Optional, Dict, Any
import pandas as pd


@dataclass
class ProfitTableRow:
    """毛利表行数据模型"""
    config: str = ""
    speed: str = ""
    name: str = ""
    price: str = "0.00"
    cost: str = "0.00"
    express: str = "30.00"
    profit: str = "0.00"
    profit_rate: str = "0.00%"
    
    def to_dict(self) -> Dict[str, str]:
        """转换为字典"""
        return {
            '配置': self.config,
            '速别': self.speed,
            '简称': self.name,
            '价格': self.price,
            '成本': self.cost,
            '快递': self.express,
            '毛利润': self.profit,
            '毛利率': self.profit_rate
        }


@dataclass
class OriginalDataRow:
    """原始数据行模型"""
    product_id: Optional[str] = None
    spec_id: Optional[str] = None
    category: str = ""
    name: str = ""
    price: str = "0.00"
    cost: str = "0.00"
    profit: Optional[str] = None
    profit_rate: Optional[str] = None
    
    # 提取的信息
    extracted_category: str = ""
    config: str = ""
    size: str = ""
    speed: str = ""
    color: str = ""
    
    # 更新后的信息
    modified_price: Optional[float] = None
    new_profit_rate: Optional[str] = None
    unmatched: bool = False


@dataclass
class ProcessingResult:
    """数据处理结果"""
    success: bool
    message: str
    data: Optional[pd.DataFrame] = None
    row_count: int = 0
    
    def __post_init__(self):
        if self.data is not None and self.row_count == 0:
            self.row_count = len(self.data)