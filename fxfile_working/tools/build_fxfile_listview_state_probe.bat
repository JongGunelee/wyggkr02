@echo off
setlocal

set "FX_VSDEVCMD=C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\Common7\Tools\VsDevCmd.bat"
set "FX_PROBE_OBJ=%~dp0..\obj\fxfile_listview_state_probe.obj"
set "FX_PROBE_EXE=%~dp0..\bin\fxfile_listview_state_probe.exe"

if not exist "%FX_VSDEVCMD%" goto :missing_build_tools

call "%FX_VSDEVCMD%" -no_logo -arch=x64
if errorlevel 1 exit /b %errorlevel%

cl.exe /nologo /W4 /EHsc "%~dp0fxfile_listview_state_probe.cpp" /Fo"%FX_PROBE_OBJ%" /Fe"%FX_PROBE_EXE%"
exit /b %errorlevel%

:missing_build_tools
echo Visual Studio Build Tools not found.
exit /b 1
