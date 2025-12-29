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

REM Set the path to the virtual environment
set "VENV_DIR=E:\automation\automation\.venv"

REM Set the path to the linkedin scripts
set "SCRIPT_DIR=E:\automation\automation\linkedin\profiles_data_extractor\common"

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

echo [INFO] Running main.py with Streamlit...
echo [INFO] The LinkedIn Reachout Hub should open in your default browser.
streamlit run main.py --browser.gatherUsageStats false --server.headless false

if errorlevel 1 (
    echo [ERROR] Application failed with error code %errorlevel%
    echo [INFO] Press any key to see error details...
    pause >nul
    goto :eof
)

echo [INFO] LinkedIn Reachout Hub closed.