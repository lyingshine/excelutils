import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import sys
import os
from typing import Optional

# 添加项目根目录到Python路径
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

try:
    from data_processor import DataProcessor
    from excel_exporter import ExcelExporter
except ImportError as e:
    print(f"导入错误: {e}")
    print("请确保data_processor.py和excel_exporter.py文件存在")
    raise

class MainWindow:
    def __init__(self, root):
        self.root = root
        self.root.title("毛利表生成器")
        self.root.geometry("1200x800")
        self.root.resizable(True, True)

        # 数据处理器和导出器
        self.data_processor = DataProcessor()
        self.excel_exporter = ExcelExporter()

        # 数据存储
        self.original_data = None
        self.processed_data = None
        self.profit_table_data = None
        self.updated_data = None

        self.setup_ui()

    def setup_ui(self):
        """设置用户界面"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 配置权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)
        
        # 初始化profit_tree
        self.profit_tree = None

        # 标题
        title_label = ttk.Label(main_frame, text="毛利表生成器", font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))

        # 按钮框架
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))

        # 功能按钮
        ttk.Button(button_frame, text="导入Excel数据", command=self.import_data, width=20).pack(side=tk.LEFT)
        ttk.Button(button_frame, text="导出毛利表", command=self.export_excel, width=20).pack(side=tk.LEFT, padx=(10, 0))
        ttk.Button(button_frame, text="导入更新后的毛利表", command=self.import_modified_prices, width=25).pack(side=tk.LEFT, padx=(10, 0))
        ttk.Button(button_frame, text="导出改价后原始数据", command=self.export_original_data, width=25).pack(side=tk.LEFT, padx=(10, 0))

        # 状态标签
        self.status_label = ttk.Label(button_frame, text="请导入数据文件", foreground="blue")
        self.status_label.pack(side=tk.LEFT, padx=(20, 0))

        # 创建Notebook用于标签页
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))

        # 毛利表标签页
        self.profit_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.profit_frame, text="毛利表")

        # 创建数据表格
        self.setup_data_tables()

    def setup_data_tables(self):
        """设置数据表格"""
        # 毛利表表格
        self.profit_tree = ttk.Treeview(self.profit_frame)
        profit_scrollbar_v = ttk.Scrollbar(self.profit_frame, orient="vertical", command=self.profit_tree.yview)
        profit_scrollbar_h = ttk.Scrollbar(self.profit_frame, orient="horizontal", command=self.profit_tree.xview)
        self.profit_tree.configure(yscrollcommand=profit_scrollbar_v.set, xscrollcommand=profit_scrollbar_h.set)

        self.profit_tree.grid(row=0, column=0, sticky="nsew")
        profit_scrollbar_v.grid(row=0, column=1, sticky="ns")
        profit_scrollbar_h.grid(row=1, column=0, sticky="ew")

        self.profit_frame.columnconfigure(0, weight=1)
        self.profit_frame.rowconfigure(0, weight=1)

    def import_data(self):
        """导入Excel数据"""
        from tkinter import filedialog
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )

        if file_path:
            try:
                self.original_data = self.data_processor.import_excel_data(file_path)
                self.status_label.config(text=f"数据导入成功，共{len(self.original_data)}行数据", foreground="blue")
                # 自动处理数据
                self.auto_process_data()
            except Exception as e:
                messagebox.showerror("错误", f"导入数据失败：{str(e)}")
                self.status_label.config(text="导入失败", foreground="red")

    def import_modified_prices(self):
        """导入更新后的毛利表"""
        if self.original_data is None:
            messagebox.showwarning("警告", "请先导入原始数据")
            return

        file_path = filedialog.askopenfilename(
            title="选择更新后的毛利表Excel文件",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )

        if file_path:
            try:
                # 读取更新后的毛利表
                modified_profit_table = self.data_processor.import_excel_data(file_path)

                # 更新价格
                self.updated_data = self.data_processor.update_prices(
                    self.original_data.copy(), modified_profit_table
                )

                # 显示成功消息
                self.status_label.config(
                    text="已成功导入更新后的毛利表，可以使用'导出改价后原始数据'按钮导出", 
                    foreground="green"
                )

                # 使用更新后的毛利表重新生成毛利表
                self.profit_table_data = modified_profit_table
                self.display_profit_table()

            except Exception as e:
                messagebox.showerror("错误", f"导入更新后毛利表失败：{str(e)}")
                self.status_label.config(text="价格更新失败", foreground="red")

    def auto_process_data(self):
        """自动处理数据"""
        try:
            if self.original_data is not None:
                self.processed_data = self.data_processor.process_data(self.original_data)
                self.generate_profit_table()
            else:
                messagebox.showwarning("警告", "请先导入数据")
        except Exception as e:
            messagebox.showerror("错误", f"数据处理失败：{str(e)}")
            self.status_label.config(text="处理失败", foreground="red")

    def generate_profit_table(self):
        """生成毛利表"""
        try:
            if self.processed_data is not None:
                self.profit_table_data = self.data_processor.generate_profit_table(self.processed_data)
                self.display_profit_table()
                self.status_label.config(text="毛利表生成完成，已按配置最高价升序排列", foreground="green")
            else:
                messagebox.showwarning("警告", "请先处理数据")
        except Exception as e:
            messagebox.showerror("错误", f"生成毛利表失败：{str(e)}")
            self.status_label.config(text="生成失败", foreground="red")

    def display_profit_table(self):
        """显示毛利表数据"""
        # 清除现有数据
        for item in self.profit_tree.get_children():
            self.profit_tree.delete(item)

        if self.profit_table_data is not None:
            # 设置列
            columns = list(self.profit_table_data.columns)
            self.profit_tree['columns'] = columns
            self.profit_tree['show'] = 'headings'

            # 设置列标题和宽度
            for col in columns:
                self.profit_tree.heading(col, text=col)
                if col in ['毛利率']:
                    self.profit_tree.column(col, width=100, minwidth=80)
                else:
                    self.profit_tree.column(col, width=120, minwidth=80)

            # 插入数据
            previous_category = None
            for index, row in self.profit_table_data.iterrows():
                values = []
                for col in columns:
                    if col == '配置':
                        current_category = str(row[col]) if pd.notna(row[col]) else ""
                        if current_category == previous_category:
                            values.append("")
                        else:
                            values.append(current_category)
                            previous_category = current_category
                    else:
                        cell_value = row[col]
                        if pd.isna(cell_value):
                            values.append("")
                        else:
                            values.append(str(cell_value))

                self.profit_tree.insert("", tk.END, values=values)

    def export_excel(self):
        """导出毛利表到Excel"""
        if self.profit_table_data is None:
            messagebox.showwarning("警告", "请先导入数据")
            return

        file_path = filedialog.asksaveasfilename(
            title="保存毛利表",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )

        if file_path:
            try:
                self.excel_exporter.export_profit_table(file_path, self.profit_table_data, self.original_data)
                # 已由导出器弹出“是否打开文件”对话框，这里不再重复弹窗
                self.status_label.config(text=f"毛利表导出成功：{file_path}", foreground="green")
            except Exception as e:
                messagebox.showerror("错误", f"导出失败：{str(e)}")
                self.status_label.config(text="导出失败", foreground="red")

    def export_original_data(self):
        """导出修改后的原始数据"""
        if self.original_data is None:
            messagebox.showwarning("警告", "请先导入数据")
            return

        # 使用更新后的数据（如果存在），否则使用原始数据
        if self.updated_data is not None:
            export_data = self.updated_data
        else:
            export_data = self.original_data.copy()
            messagebox.showinfo("提示", "请先导入更新后的毛利表，当前导出的数据将使用原价格")

        file_path = filedialog.asksaveasfilename(
            title="导出修改后的原始数据表",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )

        if file_path:
            try:
                self.excel_exporter.export_original_data(file_path, export_data)
                messagebox.showinfo("成功", f"修改后的原始数据表已保存到：{file_path}")
                self.status_label.config(text="原始数据导出成功", foreground="green")
            except Exception as e:
                messagebox.showerror("错误", f"导出原始数据失败：{str(e)}")
                self.status_label.config(text="原始数据导出失败", foreground="red")

    def run(self):
        """运行应用程序"""
        self.root.mainloop()