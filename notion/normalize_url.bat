@echo off
REM ============================================================================
REM NOTION URL NORMALIZER BATCH SCRIPT
REM ============================================================================
REM Description: This batch file runs the normalize_url.py script to clean URLs
REM              in the Notion database by removing unnecessary query parameters.
REM
REM Usage: Simply double-click this bat file or run it from command line.
REM ============================================================================

echo [INFO] Starting Notion URL Normalization process...

REM Set the path to the virtual environment
set "VENV_DIR=E:\automation\automation\.venv"

REM Set the path to the notion scripts
set "SCRIPT_DIR=E:\automation\automation\notion"

set "VENV_PY=%VENV_DIR%\Scripts\python.exe"
if not exist "%VENV_PY%" (
    echo [ERROR] Virtual environment interpreter not found at "%VENV_PY%"
    pause
    exit /b 1
)

echo [INFO] Changing to script directory: "%SCRIPT_DIR%"
cd /d "%SCRIPT_DIR%"
if errorlevel 1 (
    echo [ERROR] Failed to change directory.
    exit /b 1
)

echo [INFO] Running normalize_url.py...
"%VENV_PY%" normalize_url.py --days 14 --config normalize_url.json
if errorlevel 1 (
    echo [ERROR] normalize_url.py failed with error code %errorlevel%
    pause
    exit /b 1
)

echo [INFO] URL normalization completed successfully!
pause