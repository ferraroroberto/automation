@echo off
REM ============================================================================
REM NOTION NAME NORMALIZER BATCH SCRIPT
REM ============================================================================
REM Description: This batch file runs the normalize_names.py script to normalize
REM              article names in the Notion database to sentence case.
REM
REM Usage: Simply double-click this bat file or run it from command line.
REM ============================================================================

echo [INFO] Starting Notion Name Normalization process...

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

echo [INFO] Running normalize_names.py...
"%VENV_PY%" normalize_names.py --days 14 --config normalize_names.json
if errorlevel 1 (
    echo [ERROR] normalize_names.py failed with error code %errorlevel%
    pause
    exit /b 1
)

echo [INFO] Name normalization completed successfully!
pause