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

for %%I in ("%~dp0..") do set "REPO_ROOT=%%~fI"

REM Use the repo's own virtual environment (resolved relative to this file)
set "VENV_DIR=%REPO_ROOT%\.venv"

REM Resolve the script directory from this file (works from any checkout)
for %%I in ("%~dp0.") do set "SCRIPT_DIR=%%~fI"

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