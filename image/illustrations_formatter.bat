@echo off
REM ============================================================================
REM ILLUSTRATIONS FORMATTER GUI BATCH SCRIPT
REM ============================================================================
REM Description: This batch file runs the Illustrations Formatter GUI application.
REM              It provides a graphical interface for formatting images for
REM              Instagram or to fixed dimensions (1920x1080).
REM
REM Usage: Simply double-click this bat file or run it from command line.
REM ============================================================================

echo [INFO] Starting Illustrations Formatter GUI process...

for %%I in ("%~dp0..") do set "REPO_ROOT=%%~fI"

REM Use the repo's own virtual environment (resolved relative to this file)
set "VENV_DIR=%REPO_ROOT%\.venv"

REM Set the path to the illustrations formatter scripts
for %%I in ("%~dp0.") do set "SCRIPT_DIR=%%~fI"

echo [INFO] Using virtual environment: "%VENV_DIR%"
echo [INFO] Changing to script directory: "%SCRIPT_DIR%"

cd /d "%SCRIPT_DIR%"
if errorlevel 1 (
    echo [ERROR] Failed to change directory to %SCRIPT_DIR%
    exit /b 1
)

echo [INFO] Running illustrations_formatter_gui.py...
"%VENV_DIR%\Scripts\python.exe" illustrations_formatter_gui.py
if errorlevel 1 (
    echo [ERROR] illustrations_formatter_gui.py returned an error code: %errorlevel%
    echo [INFO] Check the logs above for more details.
    pause
    exit /b 1
)

echo [INFO] Illustrations Formatter GUI process completed successfully!
pause
