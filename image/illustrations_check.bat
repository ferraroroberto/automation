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

REM Read the virtual environment path from .env file
for /f "tokens=2 delims==" %%a in ('findstr "VENV_FOLDER" "E:\automation\automation\.env"') do set "VENV_DIR=%%a"
if not defined VENV_DIR (
    echo [WARNING] Could not read VENV_FOLDER from .env file. Using default path...
    set "VENV_DIR=E:\automation\automation\.venv"
)

set "SCRIPT_DIR=E:\automation\automation\image"

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
