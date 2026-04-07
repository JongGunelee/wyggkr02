@echo off
setlocal EnableDelayedExpansion
chcp 65001 >nul
cd /d "%~dp0"

echo [Automation Dashboard Launcher v3.5.7]
echo ---------------------------------------

set "OFFICIAL_PY_DIR=%LOCALAPPDATA%\Programs\Python\Python313"
set "PYTHON_EXE=%OFFICIAL_PY_DIR%\python.exe"

set "FOUND_APP_DIR="
if exist "..\automated_app\" set "FOUND_APP_DIR=..\automated_app"
if exist "automated_app\" set "FOUND_APP_DIR=automated_app"

if defined FOUND_APP_DIR echo [INFO] Assets found.

echo [*] Checking Python...
if not exist "!PYTHON_EXE!" goto :INSTALL_PY
"!PYTHON_EXE!" -V | find "3.13" >nul
if !ERRORLEVEL! NEQ 0 goto :INSTALL_PY
goto :CHECK_LIBS

:INSTALL_PY
echo [WARN] Python 3.13 missing. Recovering...
set "INSTALLER="
if defined FOUND_APP_DIR (
    if exist "!FOUND_APP_DIR!\python_installer.exe" set "INSTALLER=!FOUND_APP_DIR!\python_installer.exe"
)
if not defined INSTALLER (
    echo [INFO] Downloading Python...
    powershell -Command "Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.13.2/python-3.13.2-amd64.exe' -OutFile 'py_tmp.exe'"
    set "INSTALLER=py_tmp.exe"
)
start /wait "" "!INSTALLER!" /quiet InstallAllUsers=0 PrependPath=1
if exists "py_tmp.exe" del /q py_tmp.exe
if not exist "!PYTHON_EXE!" (
    echo [ERROR] Python installation failed.
    pause
    exit /b 1
)

:CHECK_LIBS
echo [*] Checking Libraries...
"!PYTHON_EXE!" -c "import win32com, fitz, openpyxl, psutil, streamlit, requests" >nul 2>&1
if !ERRORLEVEL! EQU 0 goto :START_SERVER

echo [WARN] Libraries missing. Recovering...
set "PKGS=pywin32 PyMuPDF openpyxl psutil requests streamlit"
if defined FOUND_APP_DIR (
    "!PYTHON_EXE!" -m pip install --no-index --find-links="!FOUND_APP_DIR!\packages" %PKGS%
) else (
    "!PYTHON_EXE!" -m pip install %PKGS%
)

:START_SERVER
echo [*] Starting Server...
start "Dashboard_Server" /min "!PYTHON_EXE!" "run_dashboard.py"
ping 127.0.0.1 -n 5 >nul

echo [*] Opening Dashboard...
"!PYTHON_EXE!" -c "import webbrowser; webbrowser.open('http://127.0.0.1:8501')"
exit /b 0
