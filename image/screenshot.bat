@echo off
chcp 65001 >nul

set "PROJECT_DIR=E:\automation\automation"
set "SCRIPT_PATH=E:\automation\automation\image\screenshot.py"
set "VENV_ACTIVATE=%PROJECT_DIR%\.venv\Scripts\activate.bat"

echo [INFO] Changing to project directory: "%PROJECT_DIR%\home\image"
cd /d "E:\automation\automation"
if errorlevel 1 (
    echo [ERROR] Failed to change directory.
    exit /b 1
)

REM Check for virtual environment
if exist "%VENV_ACTIVATE%" (
    echo [INFO] Activating virtual environment: "%VENV_ACTIVATE%"
    call "%VENV_ACTIVATE%"
    if errorlevel 1 (
        echo [ERROR] Failed to activate virtual environment.
        exit /b 1
    )
) else (
    echo [INFO] No virtual environment found. Using system Python.
)

echo [INFO] Running script: "%SCRIPT_PATH%" with 3
python "%SCRIPT_PATH%" 3
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