import pandas as pd
from data_processor import DataProcessor
from excel_exporter import ExcelExporter
import sys
import os

def main():
    """命令行版本：直接生成毛利表Excel文件"""
    try:
        # 初始化处理器和导出器
        processor = DataProcessor()
        exporter = ExcelExporter()
        
        # 导入数据
        input_file = "导入数据.xlsx"
        if not os.path.exists(input_file):
            print(f"错误：找不到输入文件 {input_file}")
            return
        
        print("正在导入数据...")
        original_data = processor.import_excel_data(input_file)
        
        # 处理数据
        print("正在处理数据...")
        processed_data = processor.process_data(original_data)
        
        # 生成毛利表
        print("正在生成毛利表...")
        profit_table = processor.generate_profit_table(processed_data)
        
        # 导出毛利表
        output_file = "毛利表.xlsx"
        print(f"正在导出毛利表到 {output_file}...")
        
        # 使用导出器导出
        exporter.export_profit_table(output_file, profit_table, original_data)
        
        print(f"✓ 毛利表已成功生成并保存到: {output_file}")
        print(f"✓ 共生成 {len(profit_table)} 行数据")
        
        # 显示毛利表内容
        print("\n毛利表内容预览:")
        print(profit_table.to_string())
        
    except Exception as e:
        print(f"错误: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()