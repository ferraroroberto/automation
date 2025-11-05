@echo off
REM ============================================================================
REM LINKEDIN PROFILE OPENER BATCH SCRIPT
REM ============================================================================
REM Description: This batch file runs the LinkedIn profile opener script.
REM              It opens LinkedIn profiles from Excel data with filtering options.
REM
REM Usage: Simply double-click this bat file or run it from command line.
REM ============================================================================

echo [INFO] Starting LinkedIn Profile Opener process...

REM Read the virtual environment path from .env file
for /f "tokens=2 delims==" %%a in ('findstr "VENV_FOLDER" "E:\automation\automation\.env"') do set "VENV_DIR=%%a"
if not defined VENV_DIR (
    echo [WARNING] Could not read VENV_FOLDER from .env file. Using default path...
    set "VENV_DIR=E:\automation\automation\.venv"
)

REM Set the path to the LinkedIn scripts
set "SCRIPT_DIR=E:\automation\automation\linkedin\open_profiles"

echo [INFO] Using virtual environment: "%VENV_DIR%"
echo [INFO] Changing to script directory: "%SCRIPT_DIR%"

cd /d "%SCRIPT_DIR%"
if errorlevel 1 (
    echo [ERROR] Failed to change directory to %SCRIPT_DIR%
    exit /b 1
)

echo [INFO] Running linkedin_open.py...
"%VENV_DIR%\Scripts\python.exe" linkedin_open.py
if errorlevel 1 (
    echo [ERROR] linkedin_open.py returned an error code: %errorlevel%
    echo [INFO] Check the logs above for more details.
    pause
    exit /b 1
)

echo [INFO] LinkedIn Profile Opener process completed successfully!
pause
