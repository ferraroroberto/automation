@echo off
REM ============================================================================
REM GPU RECOVERY BATCH SCRIPT
REM ============================================================================
REM Description: Diagnose and recover an NVIDIA GPU that Windows marked
REM              "disabled" (Device Manager Code 22). Falls back to the
REM              Microsoft Basic Render Driver when that happens, killing
REM              NVENC acceleration for the video tools in this folder.
REM              See gpu_recovery.md for the full story and the fix.
REM
REM Usage: Double-click this bat file (it requests admin), or run from a
REM        terminal. Enabling a device requires administrator rights.
REM ============================================================================

REM --- Self-elevate: re-launch this script with admin rights if needed --------
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo [INFO] Requesting administrator privileges...
    "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -Command "Start-Process -FilePath '%~f0' -ArgumentList '%*' -Verb RunAs"
    exit /b
)

echo [INFO] Starting GPU Recovery (administrator)...

REM Read the virtual environment path from .env file
for /f "tokens=2 delims==" %%a in ('findstr "VENV_FOLDER" "E:\automation\automation\.env"') do set "VENV_DIR=%%a"
if not defined VENV_DIR (
    echo [WARNING] Could not read VENV_FOLDER from .env file. Using default path...
    set "VENV_DIR=E:\automation\automation\.venv"
)

REM Set the path to the video scripts
set "SCRIPT_DIR=E:\automation\automation\video"

echo [INFO] Using virtual environment: "%VENV_DIR%"
echo [INFO] Changing to script directory: "%SCRIPT_DIR%"

cd /d "%SCRIPT_DIR%"
if errorlevel 1 (
    echo [ERROR] Failed to change directory to %SCRIPT_DIR%
    pause
    exit /b 1
)

echo [INFO] Running gpu_recovery.py %*...
"%VENV_DIR%\Scripts\python.exe" gpu_recovery.py %*
if errorlevel 1 (
    echo [ERROR] gpu_recovery.py returned an error code: %errorlevel%
    echo [INFO] Check the output above for details.
    pause
    exit /b 1
)

echo [INFO] GPU Recovery completed.
pause
