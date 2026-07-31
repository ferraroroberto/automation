@echo off
REM ============================================================================
REM NOTION NEWSLETTER BUILDER BATCH SCRIPT
REM ============================================================================
REM Description: This batch file runs the build_newsletter.py script to build
REM              a newsletter from Notion database articles.
REM
REM Usage: Simply double-click this bat file or run it from command line.
REM        You will be prompted to enter the newsletter number.
REM ============================================================================

echo [INFO] Starting Notion Newsletter Builder process...

REM Set the path to the virtual environment
set "VENV_DIR=E:\automation\automation\.venv"

REM Set the path to the notion scripts
set "SCRIPT_DIR=E:\automation\automation\notion"

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

REM Prompt for newsletter number
set /p NEWSLETTER_NUM="Enter newsletter number (e.g., 057): "
if "%NEWSLETTER_NUM%"=="" (
    echo [ERROR] Newsletter number is required.
    pause
    exit /b 1
)

echo [INFO] Running build_newsletter.py for newsletter %NEWSLETTER_NUM%...
"%VENV_PY%" build_newsletter.py --newsletter "%NEWSLETTER_NUM%"
if errorlevel 1 (
    echo [ERROR] build_newsletter.py failed with error code %errorlevel%
    pause
    exit /b 1
)

echo [INFO] Newsletter build completed successfully!
pause