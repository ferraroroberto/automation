@echo off
REM Toggles "luz despacho": if ON turns OFF, if OFF turns ON.
REM Runs from this script's own folder with the repo virtual environment's interpreter.
cd /d "%~dp0"
set "VENV_PY=%~dp0..\.venv\Scripts\python.exe"
if not exist "%VENV_PY%" (
    echo [ERROR] Virtual environment interpreter not found at "%VENV_PY%"
    pause
    exit /b 1
)

"%VENV_PY%" light_control.py switch --light "luz despacho"

pause
