@echo off
REM ============================================================================
REM CSV STRUCTURE COMPARISON AND PROCESSING BATCH SCRIPT
REM ============================================================================
REM Description: This batch file runs the CSV structure comparison script.
REM              It compares unified_payments_all.csv with unified_payments_all_old.csv
REM              and generates a detailed report of any structural changes.
REM              Additionally, it processes the new CSV file by filtering columns
REM              and dates (default: last closed quarter), then saves to Excel.
REM
REM Usage: Simply double-click this bat file or run it from command line.
REM ============================================================================

echo [INFO] Starting CSV Structure Comparison process...

REM Read the virtual environment path from .env file
for /f "tokens=2 delims==" %%a in ('findstr "VENV_FOLDER" "E:\automation\automation\.env"') do set "VENV_DIR=%%a"
if not defined VENV_DIR (
    echo [WARNING] Could not read VENV_FOLDER from .env file. Using default path...
    set "VENV_DIR=E:\automation\automation\.venv"
)

REM Set the path to the CSV comparison scripts
set "SCRIPT_DIR=E:\automation\automation\excel\accounting_stripe"

echo [INFO] Using virtual environment: "%VENV_DIR%"
echo [INFO] Changing to script directory: "%SCRIPT_DIR%"

cd /d "%SCRIPT_DIR%"
if errorlevel 1 (
    echo [ERROR] Failed to change directory to %SCRIPT_DIR%
    exit /b 1
)

echo [INFO] Running csv_structure_compare.py...
"%VENV_DIR%\Scripts\python.exe" csv_structure_compare.py
if errorlevel 1 (
    echo [ERROR] csv_structure_compare.py returned an error code: %errorlevel%
    echo [INFO] Check the logs above for more details.
    pause
    exit /b 1
)

echo [INFO] CSV Structure Comparison process completed successfully!
pause