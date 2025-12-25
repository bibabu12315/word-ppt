@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ==========================================
echo      正在启动 Word转PPT 助手...
echo ==========================================

if not exist ".venv\Scripts\python.exe" (
    echo [错误] 未找到虚拟环境，请先在 VS Code 中运行环境配置。
    pause
    exit /b
)

".venv\Scripts\python.exe" -m streamlit run app.py

if %errorlevel% neq 0 (
    echo.
    echo [错误] 程序异常退出。
    pause
)
