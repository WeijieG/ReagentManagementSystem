@echo off
setlocal EnableExtensions EnableDelayedExpansion

rem ============================================================
rem  Reagent Management System - one-click packaging tool
rem
rem  This is only a wrapper around tools\build_installer.py.
rem  It locates Python 3.8 x64, makes sure the build dependencies
rem  exist, optionally runs the unit tests, then starts the build.
rem
rem  Run with no arguments for an interactive menu.
rem ============================================================

pushd "%~dp0" || exit /b 1

set "INTERACTIVE=0"
set "HAS_ARGS=0"
set "SKIP_TESTS=0"
set "FORCE_PAUSE=0"
set "NO_INSTALLER=0"
set "TASK=build"
set "BUILD_ARGS="
set "CODE=0"

rem ------------------------------------------------------------
rem  Parse wrapper-only flags; everything else is passed through
rem  to tools\build_installer.py.
rem ------------------------------------------------------------
:parse_args
if "%~1"=="" goto after_parse
set "HAS_ARGS=1"
if /i "%~1"=="--no-test" (
    set "SKIP_TESTS=1"
    shift
    goto parse_args
)
if /i "%~1"=="--pause" (
    set "FORCE_PAUSE=1"
    shift
    goto parse_args
)
if /i "%~1"=="-h"     goto usage
if /i "%~1"=="--help" goto usage
if /i "%~1"=="--no-installer" set "NO_INSTALLER=1"
set "BUILD_ARGS=%BUILD_ARGS% %1"
shift
goto parse_args

:after_parse
if "%HAS_ARGS%"=="0" goto menu
goto run


rem ------------------------------------------------------------
rem  Interactive menu
rem ------------------------------------------------------------
:menu
echo.
echo ============================================================
echo   Reagent Management System - packaging tool
echo ============================================================
echo   [1] Full build: tests + exe + installer
echo   [2] Build exe only (no installer)
echo   [3] Rebuild installer only (reuse existing exe)
echo   [4] Full build + UI smoke screenshot
echo   [5] Run unit tests only
echo   [0] Exit
echo.
set "CHOICE="
set /p "CHOICE=Select [1-5, 0]: "
if not defined CHOICE goto menu

if "%CHOICE%"=="0" goto success

if "%CHOICE%"=="1" (
    set "INTERACTIVE=1"
    call :ask_version
    goto run
)
if "%CHOICE%"=="2" (
    set "INTERACTIVE=1"
    set "NO_INSTALLER=1"
    call :ask_version
    set "BUILD_ARGS=!BUILD_ARGS! --no-installer"
    goto run
)
if "%CHOICE%"=="3" (
    set "INTERACTIVE=1"
    set "BUILD_ARGS=!BUILD_ARGS! --skip-build"
    goto run
)
if "%CHOICE%"=="4" (
    set "INTERACTIVE=1"
    call :ask_version
    set "BUILD_ARGS=!BUILD_ARGS! --smoke"
    goto run
)
if "%CHOICE%"=="5" (
    set "INTERACTIVE=1"
    set "TASK=tests"
    call :find_python
    if errorlevel 1 goto fail
    call :ensure_deps
    if errorlevel 1 goto fail
    call :run_tests
    set "CODE=%ERRORLEVEL%"
    goto finish
)

echo Invalid choice.
goto menu

:ask_version
set "VER="
set /p "VER=Version override (Enter = keep config.ini version): "
if defined VER set "BUILD_ARGS=!BUILD_ARGS! --version !VER!"
exit /b 0


rem ------------------------------------------------------------
rem  Main build path
rem ------------------------------------------------------------
:run
call :find_python
if errorlevel 1 goto fail

call :ensure_deps
if errorlevel 1 goto fail

if "%SKIP_TESTS%"=="1" (
    echo.
    echo === Skipping unit tests ^(--no-test^) ===
) else (
    call :run_tests
    if errorlevel 1 goto fail
)

echo.
echo === Building: %BUILD_ARGS% ===
%BUILD_PY% tools\build_installer.py %BUILD_ARGS%
set "CODE=%ERRORLEVEL%"
goto finish


rem ------------------------------------------------------------
rem  Subroutines
rem ------------------------------------------------------------

:find_python
rem Allow an explicit interpreter via PYTHON38. Quote it because
rem the path may contain spaces.
set "BUILD_PY="
if defined PYTHON38 set "BUILD_PY="%PYTHON38%""

if not defined BUILD_PY (
    py -3.8 -c "import sys" >nul 2>&1
    if not errorlevel 1 set "BUILD_PY=py -3.8"
)

if not defined BUILD_PY (
    python -c "import sys; raise SystemExit(0 if sys.version_info[:2]==(3,8) else 1)" >nul 2>&1
    if not errorlevel 1 set "BUILD_PY=python"
)

if not defined BUILD_PY (
    echo.
    echo [ERROR] Python 3.8.x was not found.
    echo.
    echo         This project must be packaged with Python 3.8 64-bit:
    echo         it is the last Python that runs on Windows 7.
    echo.
    echo         Install Python 3.8.10 64-bit, or set PYTHON38 to the
    echo         full path of a 3.8 python.exe before running this script.
    exit /b 1
)

rem Must be 64-bit.
%BUILD_PY% -c "import sys; raise SystemExit(0 if sys.maxsize > 2**32 else 1)" >nul 2>&1
if errorlevel 1 (
    echo [ERROR] The selected Python is 32-bit. A 64-bit Python 3.8 is required.
    exit /b 1
)

set "PYVER="
for /f "delims=" %%V in ('%BUILD_PY% -V 2^>^&1') do set "PYVER=%%V"
echo [env] %PYVER%  ^(%BUILD_PY%^)
exit /b 0


:ensure_deps
%BUILD_PY% -c "import PyInstaller, PyQt5.QtWebEngineWidgets, openpyxl, requests" >nul 2>&1
if not errorlevel 1 (
    echo [env] Build dependencies OK
    exit /b 0
)

echo [env] Build dependencies missing; installing from requirements-build.txt ...
%BUILD_PY% -m pip install -r requirements-build.txt
if errorlevel 1 (
    echo.
    echo [ERROR] pip install failed.
    echo         Check the network connection, or configure a pip mirror:
    echo         python -m pip config set global.index-url https://pypi.tuna.tsinghua.edu.cn/simple
    exit /b 1
)
exit /b 0


:run_tests
echo.
echo === Running unit tests ===
%BUILD_PY% -m unittest discover -s tests -q
exit /b %ERRORLEVEL%


rem ------------------------------------------------------------
rem  Finish
rem ------------------------------------------------------------
:success
set "CODE=0"

:finish
echo.
if not "%CODE%"=="0" (
    echo [FAILED] Build exited with code %CODE%.
    echo          Common causes:
    echo            - Inno Setup ^(ISCC.exe^) not installed
    echo            - unit tests failing
    echo            - missing build dependencies
    echo            - version or production-config checks in build_installer.py
) else (
    if "%TASK%"=="tests" (
        echo [DONE] Unit tests passed.
    ) else (
        echo [DONE] Build completed.
        if "%NO_INSTALLER%"=="0" (
            if exist "Output\*.exe" (
                set "NEWEST="
                for /f "delims=" %%F in ('dir /b /o-d "Output\*.exe" 2^>nul') do (
                    if not defined NEWEST set "NEWEST=%%F"
                )
                if defined NEWEST (
                    echo   Installer output:
                    echo     %CD%\Output\!NEWEST!
                )
            )
        )
        if exist "Dist\ReagentManagementSystem" (
            echo   App folder:
            echo     %CD%\Dist\ReagentManagementSystem
        )
    )
)

if "%INTERACTIVE%"=="1" pause
if "%FORCE_PAUSE%"=="1" pause
popd
exit /b %CODE%


:fail
set "CODE=1"
goto finish


:usage
echo.
echo Usage:
echo   build.bat                          interactive menu
echo   build.bat [wrapper flags] [build args]
echo.
echo Wrapper flags:
echo   --no-test    skip the unit tests
echo   --pause      pause before exit (useful from a shortcut)
echo.
echo Build arguments passed to tools\build_installer.py:
echo   --version X.Y.Z   set the version in config.ini, then build
echo   --skip-build      reuse the existing Dist folder, rebuild installer only
echo   --no-installer    build Dist only, do not compile the installer
echo   --smoke           launch the built exe once and take a screenshot
echo.
echo Environment:
echo   PYTHON38          full path to a Python 3.8 python.exe
echo.
popd
exit /b 0
