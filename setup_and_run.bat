@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

set "PY_EXE=.venv\Scripts\python.exe"
if exist "%PY_EXE%" goto check_deps

set "BASE_PY="
where py >nul 2>nul
if not errorlevel 1 set "BASE_PY=py -3"
if not defined BASE_PY (
    where python >nul 2>nul
    if not errorlevel 1 set "BASE_PY=python"
)
if not defined BASE_PY goto no_python

echo [XTF] 正在创建虚拟环境 .venv ...
%BASE_PY% -m venv .venv
if errorlevel 1 goto venv_failed

:check_deps
if not exist "%PY_EXE%" goto no_python
if exist ".venv\.xtf_deps_ok" goto launch

echo [XTF] 首次运行，正在安装依赖（需要网络，可能需要几分钟）...
"%PY_EXE%" -m pip install --disable-pip-version-check -r requirements.txt
if errorlevel 1 goto pip_failed
type nul > ".venv\.xtf_deps_ok"

:launch
if "%XTF_SETUP_CHECK%"=="1" (
    echo [XTF] setup check passed
    exit /b 0
)
echo [XTF] 正在启动界面...
if exist ".venv\Scripts\pythonw.exe" (
    start "" ".venv\Scripts\pythonw.exe" "gui\app.py"
) else (
    start "" "%PY_EXE%" "gui\app.py"
)
exit /b 0

:no_python
echo [XTF] 未找到 Python。
echo        请先安装 Python 3.10+，安装时勾选：
echo          - Add python.exe to PATH
echo          - tcl/tk and IDLE
echo        下载地址：https://www.python.org/downloads/
pause
exit /b 1

:venv_failed
echo [XTF] 创建虚拟环境失败。请确认 Python 安装完整，并允许写入当前目录。
pause
exit /b 1

:pip_failed
echo [XTF] 依赖安装失败。请检查网络，或手动执行：
echo        .venv\Scripts\python.exe -m pip install -r requirements.txt
pause
exit /b 1
