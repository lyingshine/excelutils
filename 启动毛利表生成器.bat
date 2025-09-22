@echo off
echo 启动毛利表生成器...
echo.

REM 检查Python是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误: 未找到Python，请先安装Python
    pause
    exit /b 1
)

REM 检查依赖包
echo 检查依赖包...
python -c "import pandas, openpyxl" >nul 2>&1
if errorlevel 1 (
    echo 正在安装依赖包...
    pip install pandas openpyxl
)

REM 启动应用程序
echo 启动应用程序...
python main.py

if errorlevel 1 (
    echo.
    echo 程序运行出错，请检查错误信息
    pause
)
