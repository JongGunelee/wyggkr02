@echo off
setlocal

set "FX_VSDEVCMD=C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\Common7\Tools\VsDevCmd.bat"
set "FX_CAPTURE_OBJ=%~dp0..\obj\fxfile_window_capture.obj"
set "FX_CAPTURE_EXE=%~dp0..\bin\fxfile_window_capture.exe"

if not exist "%FX_VSDEVCMD%" goto :missing_build_tools

call "%FX_VSDEVCMD%" -no_logo -arch=x64
if errorlevel 1 exit /b %errorlevel%

cl.exe /nologo /W4 /EHsc "%~dp0fxfile_window_capture.cpp" /Fo"%FX_CAPTURE_OBJ%" /Fe"%FX_CAPTURE_EXE%"
exit /b %errorlevel%

:missing_build_tools
echo Visual Studio Build Tools not found.
exit /b 1
