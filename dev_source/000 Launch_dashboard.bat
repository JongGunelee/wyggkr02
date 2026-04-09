@echo off
setlocal EnableDelayedExpansion
chcp 65001 >nul
cd /d "%~dp0"

echo [Automation Dashboard Launcher v4.0.0]
echo ---------------------------------------

set "OFFICIAL_PY_DIR=%LOCALAPPDATA%\Programs\Python\Python313"
set "PYTHON_EXE=%OFFICIAL_PY_DIR%\python.exe"

echo [*] Checking Python...
if not exist "!PYTHON_EXE!" goto :INSTALL_PY
"!PYTHON_EXE!" -V | find "3.13" >nul
if !ERRORLEVEL! NEQ 0 goto :INSTALL_PY
goto :START_SERVER

:INSTALL_PY
echo [WARN] Python 3.13 missing. Downloading official installer...
powershell -Command "Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.13.2/python-3.13.2-amd64.exe' -OutFile 'py_tmp.exe'"
start /wait "" "py_tmp.exe" /quiet InstallAllUsers=0 PrependPath=0 Include_pip=1 Include_launcher=0 Shortcuts=0
if exist "py_tmp.exe" del /q "py_tmp.exe"
if not exist "!PYTHON_EXE!" (
    echo [ERROR] Python installation failed.
    pause
    exit /b 1
)

:START_SERVER
echo [*] Starting Server...
start "Dashboard_Server" /min "!PYTHON_EXE!" "%~dp0run_dashboard.py"
ping 127.0.0.1 -n 5 >nul

echo [*] Opening Dashboard...
"!PYTHON_EXE!" -c "import webbrowser; webbrowser.open('http://127.0.0.1:8501')"
exit /b 0
