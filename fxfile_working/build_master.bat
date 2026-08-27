@echo off
setlocal
chcp 65001 >nul
set "PATH=C:\Program Files\CMake\bin;%PATH%"
set "MSBUILDDISABLENODEREUSE=1"
set "PROJECT_DIR=%~dp0"
set "PROJECT_SUBST_TARGET=%~dp0"
if "%PROJECT_SUBST_TARGET:~-1%"=="\" set "PROJECT_SUBST_TARGET=%PROJECT_SUBST_TARGET:~0,-1%"
set "SUBST_PROBE_NAME=.fxfile_subst_probe_%RANDOM%_%RANDOM%_%RANDOM%.tmp"
set "SUBST_PROBE_PROJECT=%PROJECT_DIR%%SUBST_PROBE_NAME%"
set "SUBST_PROBE_Z=Z:\%SUBST_PROBE_NAME%"
set "BUILD_EXIT=1"
set "MAPPED_Z=0"

set "STORAGE_OVERRIDE_ARGS="
if /I "%FXFILE_STORAGE_PREFLIGHT%"=="PASS" (
    if defined FXFILE_LOW_C_OVERRIDE (
        echo [ERROR] Normal storage contract cannot inherit low-C override variables.
        exit /b 1
    )
) else if /I "%FXFILE_STORAGE_PREFLIGHT%"=="PASS_LOW_C_D_TEMP" (
    if /I not "%FXFILE_LOW_C_OVERRIDE%"=="APPROVED" (
        echo [ERROR] Low-C storage contract is missing FXFILE_LOW_C_OVERRIDE=APPROVED.
        exit /b 1
    )
    if /I not "%FXFILE_LOW_C_APPROVAL%"=="I_ACCEPT_LOW_SYSTEM_DRIVE_RISK" (
        echo [ERROR] Low-C storage contract has no exact risk acknowledgement.
        exit /b 1
    )
    if not defined FXFILE_LOW_C_INITIAL_BYTES (
        echo [ERROR] Low-C storage contract is missing the initial C: checkpoint.
        exit /b 1
    )
    if /I not "%FXFILE_LOW_C_WRITE_BUDGET_BYTES%"=="1073741824" (
        echo [ERROR] Low-C storage contract requires the exact 1 GiB incidental-write budget.
        exit /b 1
    )
    set "STORAGE_OVERRIDE_ARGS=-AllowLowSystemDriveWithDTemp -LowSystemDriveApproval I_ACCEPT_LOW_SYSTEM_DRIVE_RISK -InitialSystemFreeBytes %FXFILE_LOW_C_INITIAL_BYTES% -AllowedSystemDriveDecreaseBytes %FXFILE_LOW_C_WRITE_BUDGET_BYTES%"
) else (
    echo [ERROR] Storage preflight was not completed. Direct build_master.bat execution is blocked.
    echo [ERROR] Run build_deploy_all.bat after the documented storage checks.
    exit /b 1
)
if not defined FXFILE_BUILD_TEMP (
    echo [ERROR] FXFILE_BUILD_TEMP is not defined by the unified build workflow.
    exit /b 1
)
if /I not "%TEMP%"=="%FXFILE_BUILD_TEMP%" (
    echo [ERROR] TEMP does not match the validated FxFile build TEMP directory.
    exit /b 1
)
if /I not "%TMP%"=="%FXFILE_BUILD_TEMP%" (
    echo [ERROR] TMP does not match the validated FxFile build TEMP directory.
    exit /b 1
)
if not exist "%FXFILE_BUILD_TEMP%\" (
    echo [ERROR] Validated FxFile build TEMP directory does not exist: %FXFILE_BUILD_TEMP%
    exit /b 1
)

set "STORAGE_CHECK=%PROJECT_DIR%tools\Assert-BuildStorage.ps1"
if not exist "%STORAGE_CHECK%" (
    echo [ERROR] Build storage assertion tool is missing: %STORAGE_CHECK%
    exit /b 1
)
where pwsh.exe >nul 2>&1
if %errorlevel% equ 0 (
    pwsh.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%STORAGE_CHECK%" -ProjectRoot "%PROJECT_DIR%." -BuildTempRoot "%FXFILE_BUILD_TEMP%" %STORAGE_OVERRIDE_ARGS%
) else (
    powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%STORAGE_CHECK%" -ProjectRoot "%PROJECT_DIR%." -BuildTempRoot "%FXFILE_BUILD_TEMP%" %STORAGE_OVERRIDE_ARGS%
)
if errorlevel 1 (
    echo [ERROR] Independent build storage assertion failed.
    exit /b 1
)

set "Z_EXISTING_MAPPING="
for /f "delims=" %%M in ('subst 2^>nul ^| findstr /B /I "Z:"') do set "Z_EXISTING_MAPPING=%%M"
if defined Z_EXISTING_MAPPING (
    echo [ERROR] Z: drive is already in use. The build will not overwrite an existing drive mapping.
    echo [ERROR] Close the program using Z: or remove only a stale FxFile SUBST mapping, then retry.
    subst 2>nul | findstr /B /I "Z:"
    exit /b 1
)
if exist Z:\ (
    echo [ERROR] Z: is an existing accessible volume. The build will not overwrite it.
    exit /b 1
)

if exist "%SUBST_PROBE_PROJECT%" (
    echo [ERROR] Refusing to overwrite an existing SUBST identity probe: %SUBST_PROBE_PROJECT%
    exit /b 1
)
> "%SUBST_PROBE_PROJECT%" echo FxFile SUBST identity probe %SUBST_PROBE_NAME%
if errorlevel 1 (
    echo [ERROR] Failed to create the D: project identity probe.
    exit /b 1
)

subst Z: "%PROJECT_SUBST_TARGET%"
if errorlevel 1 (
    echo [ERROR] Failed to map Z: drive to %PROJECT_SUBST_TARGET%
    del /f /q "%SUBST_PROBE_PROJECT%" >nul 2>&1
    exit /b 1
)
set "MAPPED_Z=1"
call :verify_z_mapping
if errorlevel 1 (
    echo [ERROR] Z: mapping identity verification failed after SUBST.
    subst 2>nul | findstr /B /I "Z:"
    goto :cleanup
)

:: [Hardening] 경로 무결성 검사 (ASCII 여부)
echo %PROJECT_DIR% | findstr /R "[^a-zA-Z0-9\:\._\- ]" >nul
if %errorlevel% equ 0 (
    echo [WARNING] Current path contains non-ASCII characters (Hangeul, etc.)
    echo [WARNING] Using Z: drive as a workaround, but pure ASCII path is recommended.
)

set CMAKE_ARCH_FLAG=-A x64
set BUILD_DIR=build_cmake
if "%~1"=="x32" (
    set CMAKE_ARCH_FLAG=-A Win32
    set BUILD_DIR=build_cmake_x32
    echo [INFO] Target Architecture detected: x32 ^(Win32^)
) else (
    echo [INFO] Target Architecture detected: x64
)

cd /d Z:\ >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Failed to switch to the temporary Z: project mapping.
    goto :cleanup
)
if /I not "%CD%"=="Z:\" (
    echo [ERROR] The current directory is not the verified Z: project root: %CD%
    goto :cleanup
)
if /I "%~1"=="subst-probe" (
    echo [SUCCESS] SUBST identity, drive switch, and cleanup probe reached the verified project root.
    set "BUILD_EXIT=0"
    goto :cleanup
)
echo [INFO] Build started on Z: drive...
cmake -B "%BUILD_DIR%" -S . -G "Visual Studio 17 2022" %CMAKE_ARCH_FLAG%
if errorlevel 1 (
    echo [ERROR] CMake configuration failed!
    goto :cleanup
)
cmake --build %BUILD_DIR% --config Release -- /m /nodeReuse:false
if errorlevel 1 (
    echo [ERROR] Build execution failed!
    goto :cleanup
)
echo [SUCCESS] Build completed successfully.
set "BUILD_EXIT=0"

:cleanup
if "%MAPPED_Z%"=="1" (
    cd /d "%PROJECT_DIR%" >nul 2>&1
    if errorlevel 1 (
        echo [ERROR] Failed to leave the temporary Z: project mapping before cleanup.
        set "BUILD_EXIT=1"
        goto :cleanup_done
    )
    call :verify_z_mapping
    if errorlevel 1 (
        echo [ERROR] Refusing to remove Z: because its mapping identity changed.
        subst 2>nul | findstr /B /I "Z:"
        del /f /q "%SUBST_PROBE_PROJECT%" >nul 2>&1
        set "BUILD_EXIT=1"
        goto :cleanup_done
    )
    del /f /q "%SUBST_PROBE_PROJECT%" >nul 2>&1
    if exist "%SUBST_PROBE_PROJECT%" (
        echo [ERROR] Failed to remove the temporary SUBST identity probe.
        set "BUILD_EXIT=1"
    )
    subst Z: /D >nul 2>&1
    if errorlevel 1 (
        echo [ERROR] Failed to remove the FxFile Z: SUBST mapping.
        set "BUILD_EXIT=1"
    )
    set "Z_REMAINING_MAPPING="
    for /f "delims=" %%M in ('subst 2^>nul ^| findstr /B /I "Z:"') do set "Z_REMAINING_MAPPING=%%M"
    if defined Z_REMAINING_MAPPING (
        echo [ERROR] Z: mapping remains after cleanup; the build cannot be reported as successful.
        subst 2>nul | findstr /B /I "Z:"
        set "BUILD_EXIT=1"
    )
)
if exist "%SUBST_PROBE_PROJECT%" (
    del /f /q "%SUBST_PROBE_PROJECT%" >nul 2>&1
    if exist "%SUBST_PROBE_PROJECT%" set "BUILD_EXIT=1"
)
:cleanup_done
exit /b %BUILD_EXIT%

:verify_z_mapping
if not exist "%SUBST_PROBE_PROJECT%" (
    for /f "tokens=1,2,*" %%A in ('subst 2^>nul ^| findstr /B /I "Z:"') do (
        if /I "%%C"=="%PROJECT_SUBST_TARGET%" exit /b 0
        if /I "%%C\"=="%PROJECT_SUBST_TARGET%\" exit /b 0
    )
    exit /b 1
)
if not exist "%SUBST_PROBE_Z%" exit /b 1
fc /b "%SUBST_PROBE_PROJECT%" "%SUBST_PROBE_Z%" >nul 2>&1
if %errorlevel% equ 0 exit /b 0
for /f "tokens=1,2,*" %%A in ('subst 2^>nul ^| findstr /B /I "Z:"') do (
    if /I "%%C"=="%PROJECT_SUBST_TARGET%" exit /b 0
    if /I "%%C\"=="%PROJECT_SUBST_TARGET%\" exit /b 0
)
exit /b 1
