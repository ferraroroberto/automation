@echo off
REM Gmail and Google Drive Automation Runner
REM This script runs the email tracking automation

echo.
echo ========================================
echo   Gmail and Google Drive Automation
echo ========================================
echo.

REM Run from this script's own folder: the automation resolves its config relative to CWD
cd /d "%~dp0"

REM Interpreter of the repo virtual environment (one level up), invoked directly
set "VENV_PY=%~dp0..\.venv\Scripts\python.exe"
if not exist "%VENV_PY%" (
    echo Error: virtual environment interpreter not found at "%VENV_PY%"
    echo Create it once from the repo root: py -m venv .venv
    pause
    exit /b 1
)

REM Dependencies come from the repo root requirements.txt
echo Checking dependencies...
"%VENV_PY%" -c "import google.auth, googleapiclient" 2>nul
if errorlevel 1 (
    echo Error: Google libraries missing from the venv
    echo Install once from the repo root: .venv\Scripts\python.exe -m pip install -r requirements.txt
    pause
    exit /b 1
)

REM Run the automation
echo.
echo Starting Gmail and Drive automation...
echo.
"%VENV_PY%" gmail_drive_automation.py

REM Check if script ran successfully
if errorlevel 1 (
    echo.
    echo Error: Automation failed
    pause
    exit /b 1
) else (
    echo.
    echo Automation completed successfully!
)

echo.
echo Press any key to exit...
pause >nul
