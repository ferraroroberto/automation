@echo off
REM Visible launcher for main.py - keeps CMD window open for logging
REM Runs the LinkedIn IP check for illustrations script using the virtual environment interpreter directly

echo ========================================
echo LinkedIn IP Check for Illustrations
echo ========================================
echo.

REM Change to the script directory (important for Stream Deck launch)
cd /d "%~dp0"

echo [INFO] Working directory: %CD%
echo.

REM Path to your virtual environment (adjust if different)
set VENV_DIR=..\..\.venv

REM Check if virtual environment exists
if not exist "%VENV_DIR%\Scripts\python.exe" (
    echo [WARNING] Virtual environment not found at %VENV_DIR%
    echo [INFO] Using system Python installation...
    echo.
    python main.py
) else (
    echo [INFO] Using virtual environment at %VENV_DIR%
    echo [INFO] Running script with venv Python...
    echo.
    REM Run the script with Python from the venv (no activation needed)
    "%VENV_DIR%\Scripts\python.exe" main.py
)

echo.
echo ========================================
echo IP Check Process finished
echo ========================================
echo.
echo Press any key to close this window...
pause >nul
