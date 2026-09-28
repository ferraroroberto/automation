@echo off
REM ============================================================================
REM LINKEDIN REACHOUT HUB - MAIN APPLICATION
REM ============================================================================
REM Description: This batch file runs the main Streamlit application for the
REM LinkedIn reachout hub, providing both dashboard visualization and data entry.
REM
REM Usage: Simply double-click this bat file.
REM ============================================================================

echo [INFO] Starting LinkedIn Reachout Hub...

for %%I in ("%~dp0..\..\..") do set "REPO_ROOT=%%~fI"

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

echo [INFO] Running main.py with Streamlit...
echo [INFO] The LinkedIn Reachout Hub should open in your default browser.
"%VENV_PY%" -m streamlit run main.py --browser.gatherUsageStats false --server.headless false

if errorlevel 1 (
    echo [ERROR] Application failed with error code %errorlevel%
    echo [INFO] Press any key to see error details...
    pause >nul
    goto :eof
)

echo [INFO] LinkedIn Reachout Hub closed.