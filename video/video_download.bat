@echo off
REM ============================================================================
REM VIDEO DOWNLOADER GUI BATCH SCRIPT
REM ============================================================================
REM Description: This batch file runs the Video Downloader GUI application.
REM              Supports YouTube (yt-dlp), HLS/M3U8 (ffmpeg), and direct
REM              HTTP URLs with resume. See video_download.md for examples.
REM
REM Usage: Double-click this bat file or run from command line.
REM ============================================================================

echo [INFO] Starting Video Downloader GUI...

REM Read the virtual environment path from .env file
for /f "tokens=2 delims==" %%a in ('findstr "VENV_FOLDER" "E:\automation\automation\.env"') do set "VENV_DIR=%%a"
if not defined VENV_DIR (
    echo [WARNING] Could not read VENV_FOLDER from .env file. Using default path...
    set "VENV_DIR=E:\automation\automation\.venv"
)

REM Set the path to the video download script
set "SCRIPT_DIR=E:\automation\automation\video"

echo [INFO] Using virtual environment: "%VENV_DIR%"
echo [INFO] Changing to script directory: "%SCRIPT_DIR%"

cd /d "%SCRIPT_DIR%"
if errorlevel 1 (
    echo [ERROR] Failed to change directory to %SCRIPT_DIR%
    pause
    exit /b 1
)

echo [INFO] Running video_download.py...
"%VENV_DIR%\Scripts\python.exe" video_download.py
if errorlevel 1 (
    echo [ERROR] video_download.py returned an error code: %errorlevel%
    echo [INFO] Check the logs above for more details.
    pause
    exit /b 1
)

echo [INFO] Video Downloader process completed successfully!
pause
