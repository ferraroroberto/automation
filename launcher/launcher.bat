@echo off
REM ============================================================================
REM REMOTE PROJECT LAUNCHER
REM ============================================================================
REM Starts the Flask launcher on port 5050, bound to 0.0.0.0 so it's
REM reachable from your phone over Tailscale.
REM
REM Designed to be run via Windows Task Scheduler at log on. See README.md
REM in this folder for the exact Task Scheduler setup.
REM ============================================================================

echo [INFO] Starting Remote Project Launcher...

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

echo [INFO] Launching Flask server (Ctrl+C to stop)...
python launcher.py
if errorlevel 1 (
    echo [ERROR] launcher.py exited with code %errorlevel%
    pause
    exit /b 1
)
