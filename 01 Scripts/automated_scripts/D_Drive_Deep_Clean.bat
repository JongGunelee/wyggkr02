@echo off
REM ============================================================================
REM [Antigravity & Codex] C/D Drive + Windows Deep Clean Launcher v3.2.1
REM ============================================================================
setlocal enabledelayedexpansion

REM 관리자 권한 확인 및 UAC 승격
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo [알림] 관리자 권한으로 승격을 요청합니다...
    powershell -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

cd /d "%~dp0"

REM Python 실행 파일 탐색
set "PYTHON_BIN="

REM 1. 기설치된 Python 경로 우선 점검
if exist "C:\Users\ADMIN\AppData\Local\Programs\Python\Python313\pythonw.exe" (
    set "PYTHON_BIN=C:\Users\ADMIN\AppData\Local\Programs\Python\Python313\pythonw.exe"
)

REM 2. PATH 환경변수 내 pythonw 탐색
if "%PYTHON_BIN%"=="" (
    for /f "tokens=*" %%i in ('where pythonw.exe 2^>nul') do (
        if exist "%%i" set "PYTHON_BIN=%%i"
    )
)

REM 3. PATH 환경변수 내 python 탐색
if "%PYTHON_BIN%"=="" (
    for /f "tokens=*" %%i in ('where python.exe 2^>nul') do (
        if exist "%%i" set "PYTHON_BIN=%%i"
    )
)

REM Python GUI 앱 실행
if not "%PYTHON_BIN%"=="" (
    start "" "%PYTHON_BIN%" "%~dp0D_Drive_Deep_Clean.py"
    exit /b
)

REM Python 미설치 시 PowerShell 엔진 폴백 실행
echo [경고] Python 실행 환경을 찾을 수 없어 PowerShell 엔진으로 대체 실행합니다...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0D_Drive_Deep_Clean.ps1"
pause