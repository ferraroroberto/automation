@echo off
REM ============================================================================
REM LINKEDIN PROFILES DATA EXTRACTOR BATCH SCRIPT
REM ============================================================================
REM Description: This batch file runs the LinkedIn profile data extractor
REM
REM Usage: Simply double-click this bat file or run it from command line.
REM ============================================================================

echo [INFO] Starting LinkedIn Profiles Data Extractor Module...

REM Set the path to the virtual environment
set "VENV_DIR=E:\automation\automation\.venv"

REM Set the path to the linkedin scripts
set "SCRIPT_DIR=E:\automation\automation\linkedin\profiles_data_extractor"

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

REM Main loop for continuous execution
:LOOP
echo [INFO] Running linkedin_profiles_data_extractor.py...
python linkedin_profiles_data_extractor.py
if errorlevel 1 (
    echo [ERROR] linkedin_profiles_data_extractor.py failed with error code %errorlevel%
    echo [INFO] Press any key to retry or CTRL+C to exit...
    pause >nul
    goto LOOP
)

echo [INFO] Process completed successfully!
echo [INFO] Press ENTER to continue and repeat after 5 seconds, or CTRL+C to stop...
pause >nul

echo [INFO] Waiting 5 seconds before next cycle...
timeout /t 5 /nobreak >nul

goto LOOP
