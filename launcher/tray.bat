@echo off
REM ============================================================================
REM REMOTE PROJECT LAUNCHER — TRAY GUI
REM ============================================================================
REM Starts the launcher as a system-tray app. The Flask server runs in the
REM background; double-click the tray icon to open the status + log window.
REM
REM Designed to be run via Windows Task Scheduler at log on. See README.md
REM in this folder for the exact Task Scheduler setup.
REM ============================================================================

echo [INFO] Starting Remote Project Launcher (tray mode)...

REM Path to the repo's virtual environment.
set "VENV_DIR=E:\automation\automation\.venv"

REM Path to this launcher folder.
set "SCRIPT_DIR=E:\automation\automation\launcher"

echo [INFO] Activating virtual environment at "%VENV_DIR%"...
call "%VENV_DIR%\Scripts\activate.bat"
if errorlevel 1 (
    echo [ERROR] Failed to activate virtual environment. Make sure it exists at %VENV_DIR%
    pause
    exit /b 1
)

echo [INFO] Changing to launcher directory: "%SCRIPT_DIR%"
cd /d "%SCRIPT_DIR%"
if errorlevel 1 (
    echo [ERROR] Failed to change directory.
    pause
    exit /b 1
)

echo [INFO] Starting tray app...
python tray.py
if errorlevel 1 (
    echo [ERROR] tray.py exited with code %errorlevel%
    pause
    exit /b 1
)
