@echo off
REM ─────────────────────────────────────────────────────────────────────
REM  Streamlit Capabilities Demo — Launcher
REM  Double-click this file to start the application on Windows.
REM ─────────────────────────────────────────────────────────────────────

setlocal

REM Resolve paths relative to this batch file's location
set "PROJECT_DIR=%~dp0"
set "VENV_PYTHON=%PROJECT_DIR%.venv\Scripts\python.exe"
set "APP_FILE=%PROJECT_DIR%app.py"

REM ── Check virtual environment ──────────────────────────────────────
if not exist "%VENV_PYTHON%" (
    echo.
    echo [ERROR] Virtual environment not found at: %PROJECT_DIR%.venv
    echo.
    echo Please create it first:
    echo   cd "%PROJECT_DIR%"
    echo   python -m venv .venv
    echo   .venv\Scripts\pip install -r requirements.txt
    echo.
    pause
    exit /b 1
)

REM ── Generate mock data if missing ──────────────────────────────────
if not exist "%PROJECT_DIR%data\mock_data\employees.csv" (
    echo Generating mock data...
    "%VENV_PYTHON%" "%PROJECT_DIR%scripts\generate_mock_data.py"
)

REM ── Launch Streamlit ───────────────────────────────────────────────
echo.
echo Starting Streamlit Capabilities Demo...
echo.
"%VENV_PYTHON%" -m streamlit run "%APP_FILE%" --server.headless=false

endlocal
