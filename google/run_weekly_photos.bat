@echo off
REM Weekly Photo Automation - Windows Batch Script
REM This script runs the automation with the repo virtual environment's interpreter

echo ========================================
echo    Weekly Photo Album Automation
echo ========================================
echo.

REM Run from this script's own folder: the automation resolves its config relative to CWD
cd /d "%~dp0"

REM Interpreter of the repo virtual environment (one level up), invoked directly
set "VENV_PY=%~dp0..\.venv\Scripts\python.exe"
if not exist "%VENV_PY%" (
    echo ERROR: Virtual environment interpreter not found at "%VENV_PY%"
    echo Please run the installation steps first, from the repo root:
    echo   py -m venv .venv
    echo   .venv\Scripts\python.exe -m pip install -r requirements.txt
    pause
    exit /b 1
)

echo.
echo Running photo automation...
echo.

"%VENV_PY%" weekly_photo_automation.py %*

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
