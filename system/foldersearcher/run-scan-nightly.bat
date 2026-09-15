@echo off
REM ============================================================================
REM SCHEDULED-JOB LAUNCHER - Headless folder index scan (app-launcher chain hop)
REM ============================================================================
REM Runs foldersearcher_cli.py scan in the foreground and exits with its code:
REM   0 ok, 1 index write failed, 2 usage error, 3 no roots configured,
REM   4 a configured root is missing (3 and 4 leave the old index untouched).
REM The tray app keeps system\foldersearcher.bat (pythonw, detached).

setlocal
set "SCRIPT_DIR=%~dp0"
set "SCRIPT_DIR=%SCRIPT_DIR:~0,-1%"
set "REPO_ROOT=%SCRIPT_DIR%\..\.."
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

"%PYTHON%" "%SCRIPT_DIR%\foldersearcher_cli.py" scan
exit /b %ERRORLEVEL%
