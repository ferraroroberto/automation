@echo off
REM Gmail and Google Drive Automation Runner
REM This script runs the email tracking automation

echo.
echo ========================================
echo   Gmail and Google Drive Automation
echo ========================================
echo.

REM Check if virtual environment exists
if not exist ".venv" (
    echo Creating virtual environment...
    python -m venv .venv
    if errorlevel 1 (
        echo Error: Failed to create virtual environment
        pause
        exit /b 1
    )
)

REM Activate virtual environment
echo Activating virtual environment...
call .venv\Scripts\activate.bat
if errorlevel 1 (
    echo Error: Failed to activate virtual environment
    pause
    exit /b 1
)

REM Check if requirements are installed
echo Checking dependencies...
pip show google-auth >nul 2>&1
if errorlevel 1 (
    echo Installing requirements...
    pip install -r requirements_gmail_drive.txt
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
python gmail_drive_automation.py

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

REM Deactivate virtual environment
call .venv\Scripts\deactivate.bat

echo.
echo Press any key to exit...
pause >nul