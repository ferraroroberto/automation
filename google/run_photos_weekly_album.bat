@echo off
REM ============================================================================
REM SCHEDULED-JOB LAUNCHER - Weekly Google Photos album (app-launcher Saturday job)
REM ============================================================================
REM Runs google.photos_weekly_album_job in the foreground and exits with its code:
REM   0 done / no-op / nothing to album, 1 error, 2 config error, 3 profile signed out,
REM   4 desktop locked or Chrome window hidden, 5 a Google page didn't respond.
REM Needs an unlocked, signed-in desktop: Chrome opens a visible window. Never sends mail
REM unless google\photos_weekly_album.json says "send": true.

setlocal
set "SCRIPT_DIR=%~dp0"
set "SCRIPT_DIR=%SCRIPT_DIR:~0,-1%"
set "REPO_ROOT=%SCRIPT_DIR%\.."
set "PYTHON=%REPO_ROOT%\.venv\Scripts\python.exe"
set "PYTHONUTF8=1"
set "PYTHONUNBUFFERED=1"

cd /d "%REPO_ROOT%"
if errorlevel 1 (
    echo [ERROR] Failed to change to directory: %REPO_ROOT%
    exit /b 1
)

if not exist "%PYTHON%" (
    echo [ERROR] Python not found: %PYTHON%
    exit /b 1
)

"%PYTHON%" -m google.photos_weekly_album_job
exit /b %ERRORLEVEL%
