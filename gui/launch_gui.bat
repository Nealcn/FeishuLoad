@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0.."

if exist ".venv\Scripts\pythonw.exe" (
    start "" ".venv\Scripts\pythonw.exe" "gui\app.py"
    exit /b 0
)

where pyw >nul 2>nul
if not errorlevel 1 (
    start "" pyw "gui\app.py"
    exit /b 0
)

where python >nul 2>nul
if not errorlevel 1 (
    python "gui\app.py"
    exit /b %errorlevel%
)

echo [XTF] 未找到 Python。请先安装 Python 3.10+，然后在项目目录执行：
echo        python -m pip install -r requirements.txt
pause
