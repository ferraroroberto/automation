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

echo [INFO] Activating virtual environment...
call "%VENV_DIR%\Scripts\activate.bat"
if errorlevel 1 (
    echo [ERROR] Failed to activate virtual environment. Make sure it exists at %VENV_DIR%
    echo [INFO] Attempting to continue without virtual environment activation...
)

echo [INFO] Changing to script directory: "%SCRIPT_DIR%"
cd /d "%SCRIPT_DIR%"
if errorlevel 1 (
    echo [ERROR] Failed to change directory.
    exit /b 1
)

echo [INFO] Running normalize_names.py...
python normalize_names.py --days 14 --config normalize_names.json
if errorlevel 1 (
    echo [ERROR] normalize_names.py failed with error code %errorlevel%
    pause
    exit /b 1
)

echo [INFO] Name normalization completed successfully!
pause