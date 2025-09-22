import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import sys
import os

# 添加项目根目录到Python路径
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

try:
    from gui.main_window import MainWindow
    from utils.logger import logger
except ImportError as e:
    print(f"导入错误: {e}")
    print("请确保所有模块文件都存在")
    sys.exit(1)

def main():
    """应用程序主入口"""
    try:
        logger.info("启动毛利表生成器应用程序")
        root = tk.Tk()
        app = MainWindow(root)
        app.run()
        logger.info("应用程序正常退出")
    except Exception as e:
        logger.error(f"应用程序启动失败: {e}")
        print(f"应用程序启动失败: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()