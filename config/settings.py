import os
from pathlib import Path

# 项目根目录
BASE_DIR = Path(__file__).resolve().parent.parent

# 数据目录
DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(exist_ok=True)

# 日志配置
LOG_CONFIG = {
    "level": "INFO",
    "format": "%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    "filename": BASE_DIR / "logs" / "app.log"
}

# Excel 导出配置
EXCEL_CONFIG = {
    "default_sheet_name": "毛利表",
    "backup_sheet_name": "原始数据",
    "max_rows_per_sheet": 1000000,
    "date_format": "yyyy-mm-dd",
    "currency_format": '"¥"#,##0.00'
}

# 数据处理配置
DATA_PROCESSING_CONFIG = {
    "size_patterns": [r'20寸', r'22寸', r'24寸', r'26寸', r'27\.5寸', r'28寸', r'29寸'],
    "speed_pattern": r'(\d+速)',
    "color_patterns": [r'黑', r'白', r'红', r'蓝', r'绿', r'黄', r'灰', r'银', r'金', r'粉', r'紫', r'橙'],
    "default_size": "26寸",
    "default_speed": "21速",
    "default_color": "渐变色",
    "price_adjustments": {
        "20寸": -40,
        "22寸": -20,
        "24寸": 0,
        "26寸": 0,
        "27.5寸": 20,
        "28寸": 30,
        "29寸": 40
    }
}

# UI 配置
UI_CONFIG = {
    "window_title": "毛利表生成器",
    "window_size": "1200x800",
    "font_family": "微软雅黑",
    "font_size": {
        "title": 16,
        "normal": 11,
        "small": 10
    },
    "colors": {
        "primary": "#007bff",
        "success": "#28a745",
        "warning": "#ffc107",
        "danger": "#dc3545",
        "info": "#17a2b8"
    }
}

# 支持的Excel格式
SUPPORTED_EXCEL_FORMATS = [
    "*.xlsx",
    "*.xls"
]

# 版本信息
VERSION = "1.0.0"
AUTHOR = "ExcelUtil Team"
DESCRIPTION = "毛利表生成与数据处理工具"