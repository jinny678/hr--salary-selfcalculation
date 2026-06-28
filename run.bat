@echo off
chcp 65001 >nul
title 薪酬计算交叉验证工具

echo ============================================
echo   薪酬计算交叉验证工具 v1.0
echo   广汽五羊-本田 正式工月度薪酬核验
echo ============================================
echo.

:: 检查 Python 是否安装
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 未检测到 Python，请先安装 Python 3.8+
    echo 下载地址: https://www.python.org/downloads/
    pause
    exit /b 1
)

:: 检查并安装依赖
echo [检查] 正在检查依赖包...
pip install -r requirements.txt --quiet --break-system-packages 2>nul
if %errorlevel% neq 0 (
    pip install -r requirements.txt --quiet
)

:: 运行主程序
echo.
python main.py
echo.

:: 如果出错则暂停
if %errorlevel% neq 0 (
    echo.
    echo [错误] 程序运行异常，请检查上方错误信息。
    pause
    exit /b 1
)

echo.
echo ============================================
echo   完成！报告已生成在 output 文件夹中。
echo   按任意键退出...
echo ============================================
pause >nul