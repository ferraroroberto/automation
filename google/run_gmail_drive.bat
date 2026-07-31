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

REM Check if requirements are installed
echo Checking dependencies...
"%VENV_PY%" -m pip show google-auth >nul 2>&1
if errorlevel 1 (
    echo Installing requirements...
    "%VENV_PY%" -m pip install -r requirements_gmail_drive.txt
    if errorlevel 1 (
        echo Error: Failed to install requirements
        pause
        exit /b 1
    )
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
