@echo off
REM ============================================================================
REM NOTION EDITORIAL DATABASE — ADD MISSING DATES
REM ============================================================================
REM Description: Runs notion_databases_add_editorial.py to ensure every day in
REM              the current calendar month and the next calendar month exists
REM              as a row in the Notion editorial database (skips dates already present).
REM
REM Usage: Double-click this file or run from a command prompt.
REM ============================================================================

echo [INFO] Starting Notion editorial date sync...

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

echo [INFO] Running notion_databases_add_editorial.py...
"%VENV_PY%" notion_databases_add_editorial.py
if errorlevel 1 (
    echo [ERROR] notion_databases_add_editorial.py failed with error code %errorlevel%
    pause
    exit /b 1
)

echo [INFO] Editorial date sync completed successfully!
pause
