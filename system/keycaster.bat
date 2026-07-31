@echo off
REM Runs keycaster.py with its virtual environment's interpreter

REM Path to keycaster's dedicated virtual environment (PySimpleGUI is not a repo dependency)
set VENV_DIR=E:\onedrive\Documentos\Roberto\projects\automation\notion-automation\local\.venv_keycaster

REM Run the script with Python from the venv (no activation)
set "VENV_PY=%VENV_DIR%\Scripts\python.exe"
if not exist "%VENV_PY%" (
    echo [ERROR] Virtual environment interpreter not found at "%VENV_PY%"
    pause
    exit /b 1
)

cd /d "%~dp0"
"%VENV_PY%" keycaster.py

REM Pause so you can read any error messages
pause
