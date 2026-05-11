@echo off
chcp 65001 >nul
title 路易小姐薪资计算工具
echo 🚀 正在启动薪资计算工具...
echo.

:: 检查 Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ❌ 未找到 Python，请先安装 Python 3.8+
    echo    下载地址：https://www.python.org/downloads/windows/
    pause
    exit /b 1
)

:: 安装依赖（如果未安装）
pip show flask >nul 2>&1
if %errorlevel% neq 0 (
    echo 📦 首次运行，正在安装依赖...
    pip install flask openpyxl -q
)

echo ✅ 启动成功！
echo   浏览器打开：http://localhost:5001
echo.
echo ⏎ 按 Ctrl+C 停止服务
echo.

python app.py
pause
