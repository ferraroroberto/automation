@echo off
REM ============================================================================
REM AUDIO EXTRACTOR GUI BATCH SCRIPT
REM ============================================================================
REM Description: This batch file runs the Audio Extractor GUI application.
REM              It provides a graphical interface to extract audio tracks
REM              from selected video files.
REM
REM Usage: Double-click this bat file or run it from command line.
REM ============================================================================

echo [INFO] Starting Audio Extractor GUI process...

REM Read the virtual environment path from .env file
for /f "tokens=2 delims==" %%a in ('findstr "VENV_FOLDER" "E:\automation\automation\.env"') do set "VENV_DIR=%%a"
if not defined VENV_DIR (
    echo [WARNING] Could not read VENV_FOLDER from .env file. Using default path...
    set "VENV_DIR=E:\automation\automation\.venv"
)

REM Set the path to the audio extractor scripts
set "SCRIPT_DIR=E:\automation\automation\audio"

echo [INFO] Using virtual environment: "%VENV_DIR%"
echo [INFO] Changing to script directory: "%SCRIPT_DIR%"

cd /d "%SCRIPT_DIR%"
if errorlevel 1 (
    echo [ERROR] Failed to change directory to %SCRIPT_DIR%
    exit /b 1
)

echo [INFO] Running audio_extractor_gui.py...
"%VENV_DIR%\Scripts\python.exe" audio_extractor_gui.py
if errorlevel 1 (
    echo [ERROR] audio_extractor_gui.py returned an error code: %errorlevel%
    echo [INFO] Check the logs above for more details.
    pause
    exit /b 1
)

echo [INFO] Audio Extractor GUI process completed successfully!
pause
