@echo off
setlocal
set "SCRIPT=%~dp0tools\Build-Deploy-Verify.ps1"

where pwsh.exe >nul 2>&1
if %errorlevel% equ 0 (
    pwsh.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" %*
) else (
    powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" %*
)

exit /b %errorlevel%
