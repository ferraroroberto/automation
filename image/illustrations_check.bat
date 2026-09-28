@echo off
REM ============================================================================
REM ILLUSTRATIONS CHECK BATCH SCRIPT
REM ============================================================================
REM Description: Checks a folder for matching .afdesign and .png pairs.
REM              Reports orphan files and spare files. Terminal-only.
REM
REM Usage: illustrations_check.bat [source_folder]
REM        If source_folder is omitted, you will be prompted.
REM ============================================================================

echo [INFO] Starting Illustrations Check...

for %%I in ("%~dp0..") do set "REPO_ROOT=%%~fI"

REM Use the repo's own virtual environment (resolved relative to this file)
set "VENV_DIR=%REPO_ROOT%\.venv"

for %%I in ("%~dp0.") do set "SCRIPT_DIR=%%~fI"

echo [INFO] Using virtual environment: "%VENV_DIR%"
cd /d "%SCRIPT_DIR%"
if errorlevel 1 (
    echo [ERROR] Failed to change directory to %SCRIPT_DIR%
    exit /b 1
)

echo [INFO] Running illustrations_check.py...
"%VENV_DIR%\Scripts\python.exe" illustrations_check.py %*
if errorlevel 1 (
    echo [ERROR] illustrations_check.py returned error code: %errorlevel%
    pause
    exit /b 1
)

echo.
echo [INFO] Done.
pause
