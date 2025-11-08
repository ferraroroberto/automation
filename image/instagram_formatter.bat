@echo off
REM ============================================================================
REM INSTAGRAM FORMATTER GUI BATCH SCRIPT
REM ============================================================================
REM Description: This batch file runs the Instagram Formatter GUI application.
REM              It provides a graphical interface for formatting images for Instagram.
REM
REM Usage: Simply double-click this bat file or run it from command line.
REM ============================================================================

echo [INFO] Starting Instagram Formatter GUI process...

REM Read the virtual environment path from .env file
for /f "tokens=2 delims==" %%a in ('findstr "VENV_FOLDER" "E:\automation\automation\.env"') do set "VENV_DIR=%%a"
if not defined VENV_DIR (
    echo [WARNING] Could not read VENV_FOLDER from .env file. Using default path...
    set "VENV_DIR=E:\automation\automation\.venv"
)

REM Set the path to the Instagram formatter scripts
set "SCRIPT_DIR=E:\automation\automation\image"

echo [INFO] Using virtual environment: "%VENV_DIR%"
echo [INFO] Changing to script directory: "%SCRIPT_DIR%"

cd /d "%SCRIPT_DIR%"
if errorlevel 1 (
    echo [ERROR] Failed to change directory to %SCRIPT_DIR%
    exit /b 1
)

echo [INFO] Running instagram_formatter_gui.py...
"%VENV_DIR%\Scripts\python.exe" instagram_formatter_gui.py
if errorlevel 1 (
    echo [ERROR] instagram_formatter_gui.py returned an error code: %errorlevel%
    echo [INFO] Check the logs above for more details.
    pause
    exit /b 1
)

echo [INFO] Instagram Formatter GUI process completed successfully!
pause
