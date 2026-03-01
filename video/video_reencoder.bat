@echo off
REM ============================================================================
REM VIDEO RE-ENCODER GUI BATCH SCRIPT
REM ============================================================================
REM Description: This batch file runs the Video Re-encoder GUI application.
REM              Re-encode videos to a target file size using FFmpeg.
REM              See video_reencoder.md for examples.
REM
REM Usage: Double-click this bat file or run from command line.
REM ============================================================================

echo [INFO] Starting Video Re-encoder GUI...

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

echo [INFO] Running video_reencoder.py...
"%VENV_DIR%\Scripts\python.exe" video_reencoder.py
if errorlevel 1 (
    echo [ERROR] video_reencoder.py returned an error code: %errorlevel%
    echo [INFO] Check the logs above for more details.
    pause
    exit /b 1
)

echo [INFO] Video Re-encoder process completed successfully!
pause
