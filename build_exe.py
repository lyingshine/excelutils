"""
EXE打包脚本
使用PyInstaller将项目打包成可执行文件
"""

import os
import sys
import subprocess
import shutil
from pathlib import Path

def install_pyinstaller():
    """安装PyInstaller"""
    print("正在安装PyInstaller...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
        print("✅ PyInstaller安装成功")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ PyInstaller安装失败: {e}")
        return False

def create_spec_file():
    """创建PyInstaller配置文件"""
    spec_content = '''# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

# 收集所有Python文件
a = Analysis(
    ['main.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        ('config', 'config'),
        ('data', 'data'),
        ('logs', 'logs'),
    ],
    hiddenimports=[
        'pandas',
        'openpyxl',
        'xlrd',
        'XlsxWriter',
        'numpy',
        'tkinter',
        'tkinter.ttk',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'gui.main_window',
        'core.data_extractor',
        'core.data_filter',
        'core.profit_calculator',
        'core.price_matcher',
        'core.table_format_analyzer',
        'processors.data_processor',
        'services.data_service',
        'services.excel_service',
        'exporters.excel_exporter',
        'utils.logger',
        'models',
        'app',
        'cli',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='毛利表生成器',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # 不显示控制台窗口
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # 可以添加图标文件路径
)
'''
    
    with open('毛利表生成器.spec', 'w', encoding='utf-8') as f:
        f.write(spec_content)
    print("✅ 创建配置文件成功")

def build_exe():
    """构建EXE文件"""
    print("正在构建EXE文件...")
    try:
        # 使用spec文件构建
        subprocess.check_call([
            sys.executable, "-m", "PyInstaller", 
            "--clean", 
            "毛利表生成器.spec"
        ])
        print("✅ EXE文件构建成功")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ EXE文件构建失败: {e}")
        return False

def create_distribution():
    """创建发布包"""
    print("正在创建发布包...")
    
    # 创建发布目录
    dist_dir = Path("发布包")
    if dist_dir.exists():
        shutil.rmtree(dist_dir)
    dist_dir.mkdir()
    
    # 复制EXE文件
    exe_file = Path("dist/毛利表生成器.exe")
    if exe_file.exists():
        shutil.copy2(exe_file, dist_dir / "毛利表生成器.exe")
        print("✅ 复制EXE文件成功")
    else:
        print("❌ 找不到EXE文件")
        return False
    
    # 复制必要文件
    files_to_copy = [
        "README.md",
        "CHANGELOG.md",
        "requirements.txt"
    ]
    
    for file_name in files_to_copy:
        file_path = Path(file_name)
        if file_path.exists():
            shutil.copy2(file_path, dist_dir / file_name)
            print(f"✅ 复制 {file_name} 成功")
    
    # 创建使用说明
    usage_content = """# 毛利表生成器 使用说明

## 快速开始

1. 双击 `毛利表生成器.exe` 启动程序
2. 将原始数据文件命名为 `导入数据.xlsx` 并放在程序同目录下
3. 点击"生成毛利表"按钮
4. 修改生成的 `毛利表.xlsx` 文件中的价格
5. 再次运行程序，点击"生成改价数据"按钮

## 数据格式要求

原始数据文件应包含以下列：
- 简称：产品完整名称（如：YJ-FT芭蕾顶配24寸变速）
- 分类：产品分类
- 价格：产品价格

## 功能特点

- ✅ 动态尺寸识别：支持任何"数字+寸"格式的尺寸
- ✅ 智能格式切换：根据数据特征自动选择显示格式
- ✅ 双向数据处理：原始数据→毛利表→改价数据
- ✅ 专业Excel格式化：自动排版和格式设置

## 注意事项

1. 确保数据文件格式正确
2. 程序会在同目录下生成日志文件用于调试
3. 如遇问题，请查看日志文件或联系技术支持

## 版本信息

版本：1.0.0
更新时间：2025年9月23日
"""
    
    with open(dist_dir / "使用说明.txt", 'w', encoding='utf-8') as f:
        f.write(usage_content)
    print("✅ 创建使用说明成功")
    
    # 创建启动脚本（备用）
    bat_content = """@echo off
chcp 65001 > nul
echo 启动毛利表生成器...
"毛利表生成器.exe"
pause
"""
    
    with open(dist_dir / "启动程序.bat", 'w', encoding='gbk') as f:
        f.write(bat_content)
    print("✅ 创建启动脚本成功")
    
    print(f"✅ 发布包创建完成，位置：{dist_dir.absolute()}")
    return True

def clean_build_files():
    """清理构建文件"""
    print("正在清理构建文件...")
    
    dirs_to_remove = ["build", "dist", "__pycache__"]
    files_to_remove = ["毛利表生成器.spec"]
    
    for dir_name in dirs_to_remove:
        dir_path = Path(dir_name)
        if dir_path.exists():
            shutil.rmtree(dir_path)
            print(f"✅ 删除目录 {dir_name}")
    
    for file_name in files_to_remove:
        file_path = Path(file_name)
        if file_path.exists():
            file_path.unlink()
            print(f"✅ 删除文件 {file_name}")

def main():
    """主函数"""
    print("=" * 50)
    print("🚀 毛利表生成器 EXE打包工具")
    print("=" * 50)
    
    # 检查Python版本
    if sys.version_info < (3, 8):
        print("❌ 需要Python 3.8或更高版本")
        return False
    
    print(f"✅ Python版本: {sys.version}")
    
    # 安装PyInstaller
    if not install_pyinstaller():
        return False
    
    # 创建配置文件
    create_spec_file()
    
    # 构建EXE
    if not build_exe():
        return False
    
    # 创建发布包
    if not create_distribution():
        return False
    
    # 询问是否清理构建文件
    response = input("\n是否清理构建文件？(y/n): ").lower().strip()
    if response in ['y', 'yes', '是']:
        clean_build_files()
    
    print("\n" + "=" * 50)
    print("🎉 EXE打包完成！")
    print("📁 发布包位置：发布包/")
    print("🚀 可以直接运行：发布包/毛利表生成器.exe")
    print("=" * 50)
    
    return True

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n❌ 用户取消操作")
    except Exception as e:
        print(f"\n❌ 发生错误: {e}")
        import traceback
        traceback.print_exc()
    
    input("\n按回车键退出...")