@echo off
REM ============================================================================
REM VIDEO TRIMMER GUI BATCH SCRIPT
REM ============================================================================
REM Description: This batch file runs the Video Trimmer GUI application.
REM              Trim videos with FFmpeg (GPU accelerated if available).
REM              Supports GUI and CLI. See video_trim.md for examples.
REM
REM Usage: Double-click this bat file or run from command line.
REM ============================================================================

echo [INFO] Starting Video Trimmer GUI...

for %%I in ("%~dp0..") do set "REPO_ROOT=%%~fI"

REM Use the repo's own virtual environment (resolved relative to this file)
set "VENV_DIR=%REPO_ROOT%\.venv"

REM Resolve the script directory from this file (works from any checkout)
for %%I in ("%~dp0.") do set "SCRIPT_DIR=%%~fI"

echo [INFO] Using virtual environment: "%VENV_DIR%"
echo [INFO] Changing to script directory: "%SCRIPT_DIR%"

cd /d "%SCRIPT_DIR%"
if errorlevel 1 (
    echo [ERROR] Failed to change directory to %SCRIPT_DIR%
    pause
    exit /b 1
)

echo [INFO] Running video_trim.py...
"%VENV_DIR%\Scripts\python.exe" video_trim.py
if errorlevel 1 (
    echo [ERROR] video_trim.py returned an error code: %errorlevel%
    echo [INFO] Check the logs above for more details.
    pause
    exit /b 1
)

echo [INFO] Video Trimmer process completed successfully!
pause
