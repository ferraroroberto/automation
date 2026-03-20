@echo off
REM Runs notion_databases_add_editorial.py with the repo virtual environment.
set "SCRIPT_DIR=%~dp0"
set "VENV_DIR=%SCRIPT_DIR%..\.venv"
set "PYTHON_EXE=%VENV_DIR%\Scripts\python.exe"

if not exist "%PYTHON_EXE%" (
    echo [ERROR] Python not found at "%PYTHON_EXE%"
    pause
    exit /b 1
)

cd /d "%SCRIPT_DIR%"
echo [INFO] Using "%PYTHON_EXE%"
"%PYTHON_EXE%" notion_databases_add_editorial.py
if errorlevel 1 (
    echo [ERROR] Script failed with code %errorlevel%
    pause
    exit /b 1
)
pause
