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

for %%I in ("%~dp0..\..") do set "REPO_ROOT=%%~fI"

REM Use the repo's own virtual environment (resolved relative to this file)
set "VENV_DIR=%REPO_ROOT%\.venv"

REM Resolve the script directory from this file (works from any checkout)
for %%I in ("%~dp0.") do set "SCRIPT_DIR=%%~fI"

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
