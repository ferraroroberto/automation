@echo off
REM ============================================================================
REM LINKEDIN PROFILE SEARCH CHECKER BATCH SCRIPT
REM ============================================================================
REM Description: This batch file runs the LinkedIn profile search checker
REM which extracts profile names and page numbers from search tabs and
REM compares them with existing contacts to show missing profiles
REM
REM Usage: Simply double-click this bat file or run it from command line.
REM ============================================================================

echo [INFO] Starting LinkedIn Profile Search Checker Module...

REM Set the path to the virtual environment
set "VENV_DIR=E:\automation\automation\.venv"

REM Set the path to the linkedin scripts
set "SCRIPT_DIR=E:\automation\automation\linkedin\profiles_data_extractor\brute_force"

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

echo [INFO] Running linkedin_profile_search_checker.py...
python linkedin_profile_search_checker.py
if errorlevel 1 (
    echo [ERROR] linkedin_profile_search_checker.py failed with error code %errorlevel%
    echo [INFO] Press any key to retry or CTRL+C to exit...
    pause >nul
    goto :eof
)

echo [INFO] Process completed successfully!
echo [INFO] Press any key to exit...
pause >nul
