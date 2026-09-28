@echo off
REM ============================================================================
REM NOTION JOURNAL AUTOMATION BATCH SCRIPT
REM ============================================================================
REM Description: This batch file runs the journal_automation.py script to
REM              process Notion journal database entries into a consolidated
REM              weekly summary optimized for LLM analysis.
REM
REM Usage: Simply double-click this bat file or run it from command line.
REM        You will be prompted to enter a date offset (default is 0).
REM ============================================================================

echo [INFO] Starting Notion Journal Automation process...

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

echo [INFO] Running journal_automation.py...
"%VENV_PY%" journal_automation.py
if errorlevel 1 (
    echo [ERROR] journal_automation.py failed with error code %errorlevel%
    pause
    exit /b 1
)

echo [INFO] Journal automation completed successfully!
pause
