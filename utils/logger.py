import logging
import sys
from pathlib import Path
from config.settings import LOG_CONFIG, BASE_DIR

def setup_logger(name: str = __name__) -> logging.Logger:
    """
    设置和配置日志记录器
    
    Args:
        name: 日志记录器名称
        
    Returns:
        logging.Logger: 配置好的日志记录器
    """
    # 创建日志目录
    log_dir = BASE_DIR / "logs"
    log_dir.mkdir(exist_ok=True)
    
    # 创建日志记录器
    logger = logging.getLogger(name)
    logger.setLevel(LOG_CONFIG["level"])
    
    # 清除现有的处理器
    logger.handlers.clear()
    
    # 创建格式化器
    formatter = logging.Formatter(LOG_CONFIG["format"])
    
    # 文件处理器
    file_handler = logging.FileHandler(LOG_CONFIG["filename"], encoding='utf-8')
    file_handler.setFormatter(formatter)
    file_handler.setLevel(LOG_CONFIG["level"])
    
    # 控制台处理器
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(formatter)
    console_handler.setLevel(LOG_CONFIG["level"])
    
    # 添加处理器到日志记录器
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    
    return logger

# 创建默认日志记录器
logger = setup_logger()