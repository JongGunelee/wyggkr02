@echo off
setlocal
chcp 65001 >nul
set "PROJECT_SUBST_TARGET=%~dp0.."
for %%I in ("%PROJECT_SUBST_TARGET%") do set "PROJECT_SUBST_TARGET=%%~fI"
set "PROBE_NAME=.fxfile_subst_probe_test_%RANDOM%_%RANDOM%.tmp"
set "PROBE_PROJECT=%PROJECT_SUBST_TARGET%\%PROBE_NAME%"
set "PROBE_Z=Z:\%PROBE_NAME%"
if exist Z:\ exit /b 10
> "%PROBE_PROJECT%" echo FxFile SUBST identity probe %PROBE_NAME%
if errorlevel 1 exit /b 15
subst Z: "%PROJECT_SUBST_TARGET%"
if errorlevel 1 exit /b 11
if not exist "%PROBE_Z%" goto :fail
fc /b "%PROBE_PROJECT%" "%PROBE_Z%" >nul 2>&1
if errorlevel 1 goto :fail
cd /d Z:\
if errorlevel 1 goto :fail
if /I not "%CD%"=="Z:\" goto :fail
cd /d "%PROJECT_SUBST_TARGET%"
del /f /q "%PROBE_PROJECT%" >nul 2>&1
subst Z: /D
if errorlevel 1 exit /b 12
if exist Z:\ exit /b 13
echo PASS: UTF-8 SUBST identity, drive switch, and cleanup.
exit /b 0
:fail
cd /d "%PROJECT_SUBST_TARGET%"
del /f /q "%PROBE_PROJECT%" >nul 2>&1
subst Z: /D >nul 2>&1
exit /b 14
