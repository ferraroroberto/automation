@echo off
REM Weekly Photo Automation - Windows Batch Script
REM This script activates the virtual environment and runs the automation

echo ========================================
echo    Weekly Photo Album Automation
echo ========================================
echo.

REM Check if virtual environment exists
if not exist ".venv\Scripts\activate.bat" (
    echo ERROR: Virtual environment not found!
    echo Please run the installation steps first:
    echo   python -m venv .venv
    echo   .venv\Scripts\activate
    echo   pip install -r requirements.txt
    pause
    exit /b 1
)

REM Activate virtual environment and run script
echo Activating virtual environment...
call .venv\Scripts\activate.bat

echo.
echo Running photo automation...
echo.

python weekly_photo_automation.py %*

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ========================================
    echo    Automation completed successfully!
    echo ========================================
) else (
    echo.
    echo ========================================
    echo    ERROR: Automation failed!
    echo    Check the logs for details.
    echo ========================================
)

echo.
pause