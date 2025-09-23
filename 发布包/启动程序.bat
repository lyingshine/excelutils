@echo off
chcp 65001 > nul
echo 启动毛利表生成器...
"毛利表生成器.exe"
if errorlevel 1 (
    echo.
    echo 程序运行出现错误，请检查：
    echo 1. 是否有导入数据.xlsx文件
    echo 2. 数据格式是否正确
    echo 3. 查看日志文件获取详细信息
    echo.
    pause
) else (
    echo 程序运行完成
)