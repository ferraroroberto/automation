@echo off
REM Quick Spanish Transcription Launcher
REM Launches transcription in quick mode without GUI

cd /d "%~dp0"
REM Calculate venv path from script location
for %%I in ("%~dp0..\..") do set "VENV_FOLDER=%%~fI\.venv"

echo Starting Spanish voice transcription...
echo Press any key to stop recording early
echo.

REM Run the transcription
call "%VENV_FOLDER%\Scripts\python.exe" transcribe_voice_gui.py --quick-spanish

echo.
echo Transcription completed.
