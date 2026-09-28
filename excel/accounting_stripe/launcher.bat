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

for %%I in ("%~dp0..\..") do set "REPO_ROOT=%%~fI"

REM Use the repo's own virtual environment (resolved relative to this file)
set "VENV_DIR=%REPO_ROOT%\.venv"

REM Set the path to the CSV comparison scripts
for %%I in ("%~dp0.") do set "SCRIPT_DIR=%%~fI"

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