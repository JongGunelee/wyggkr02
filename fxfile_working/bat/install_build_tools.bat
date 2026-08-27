@echo off
setlocal
set "INSTALLER=C:\Program Files (x86)\Microsoft Visual Studio\Installer\setup.exe"
set "CONFIG=%~dp0fxfile_working\.vsconfig"
set "INSTALL_DIR=C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools"

echo ============================================================
echo [FxFile] Visual Studio Build Tools C++ Component Setup
echo ============================================================
echo.
echo Installer : "%INSTALLER%"
echo Config    : "%CONFIG%"
echo Target Dir: "%INSTALL_DIR%"
echo.
echo Launching Visual Studio Installer...
echo Please wait while components are downloading and installing...
echo.

"%INSTALLER%" modify --installPath "%INSTALL_DIR%" --config "%CONFIG%" --passive --norestart

set "CODE=%ERRORLEVEL%"
echo.
echo ============================================================
echo Installer process finished with Exit Code: %CODE%
echo ============================================================
pause
