@echo off
chcp 65001 > nul
echo ========================================
echo 🚀 毛利表生成器 EXE打包工具
echo ========================================

echo 正在运行打包脚本...
python build_exe.py

echo.
echo 打包完成！
pause