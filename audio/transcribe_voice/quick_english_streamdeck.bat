@echo off
REM Quick English Transcription Launcher for Stream Deck
REM Ensures terminal is visible and interactive for Stream Deck

cd /d "%~dp0"
REM Calculate venv path from script location
for %%I in ("%~dp0..\..") do set "VENV_FOLDER=%%~fI\.venv"

REM Force console window to be visible and on top
title "English Transcription (PRESS ANY KEY TO STOP)"
mode con: cols=80 lines=25
powershell -command "$wshell = New-Object -ComObject wscript.shell; $wshell.AppActivate('English Transcription')"

cls
echo ===========================================
echo ENGLISH VOICE TRANSCRIPTION
echo ===========================================
echo.
echo Instructions:
echo - Speak clearly into your microphone
echo - Press ANY KEY to stop recording early
echo - Recording will auto-stop after 5 minutes
echo.
echo Starting recording now...
echo.

REM Run the transcription
call "%VENV_FOLDER%\Scripts\python.exe" transcribe_voice_gui.py --quick-english

REM Show completion message
echo.
echo ===========================================
echo TRANSCRIPTION COMPLETE
echo ===========================================
echo Text has been copied to clipboard.
