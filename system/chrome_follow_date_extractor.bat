@echo off
REM Chrome Follow Date Extractor Launcher
REM Extracts "following you since" date and URL from current Chrome tab

chcp 65001 >nul

set "PROJECT_DIR=E:\automation\automation"
set "SCRIPT_PATH=%PROJECT_DIR%\system\chrome_follow_date_extractor.py"
set "VENV_DIR=%PROJECT_DIR%\.venv"

echo [INFO] Starting Chrome Follow Date Extractor...
echo.

REM Check for virtual environment
if exist "%VENV_DIR%\Scripts\python.exe" (
    echo [INFO] Using virtual environment: "%VENV_DIR%"
    "%VENV_DIR%\Scripts\python.exe" "%SCRIPT_PATH%"
) else (
    echo [WARNING] Virtual environment not found. Using system Python.
    python "%SCRIPT_PATH%"
)

set "PY_EXIT_CODE=%ERRORLEVEL%"

if not "%PY_EXIT_CODE%"=="0" (
    echo [ERROR] Script exited with code %PY_EXIT_CODE%.
    echo.
    pause
) else (
    echo.
    echo [INFO] Script completed successfully.
    pause
)

exit /b %PY_EXIT_CODE%

