@echo off
chcp 65001 >nul

set "PROJECT_DIR=E:\automation\automation"
set "SCRIPT_PATH=E:\automation\automation\image\screenshot.py"
set "VENV_PY=%PROJECT_DIR%\.venv\Scripts\python.exe"

echo [INFO] Changing to project directory: "%PROJECT_DIR%"
cd /d "%PROJECT_DIR%"
if errorlevel 1 (
    echo [ERROR] Failed to change directory.
    exit /b 1
)

REM Require the project virtual environment's interpreter
if not exist "%VENV_PY%" (
    echo [ERROR] Virtual environment interpreter not found at "%VENV_PY%".
    exit /b 1
)

echo [INFO] Running script: "%SCRIPT_PATH%" with 3
"%VENV_PY%" "%SCRIPT_PATH%" 3
set "PY_EXIT_CODE=%ERRORLEVEL%"

if not "%PY_EXIT_CODE%"=="0" (
    echo [ERROR] Python script exited with code %PY_EXIT_CODE%.
    echo.
    echo [INFO] Press any key to exit...
    pause >nul
) else (
    echo [INFO] Script executed successfully with exit code 0.
)

exit /b %PY_EXIT_CODE%