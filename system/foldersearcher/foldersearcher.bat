@echo off
REM Runs foldersearcher.py with the repo virtual environment's interpreter

REM Run from this script's own folder so foldersearcher_core.py resolves
cd /d "%~dp0"

REM Interpreter of the repo virtual environment, resolved relative to this file
set "VENV_PY=%~dp0..\..\.venv\Scripts\python.exe"
if not exist "%VENV_PY%" (
    echo [ERROR] Virtual environment interpreter not found at "%VENV_PY%"
    pause
    exit /b 1
)

REM Run the script with Python from the venv
"%VENV_PY%" foldersearcher.py

REM Pause so you can read any error messages
pause 