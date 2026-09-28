@echo off
REM Launch Screen Recorder GUI silently using the repo venv.
set "SCRIPT_DIR=%~dp0"
set "REPO_ROOT=%SCRIPT_DIR%.."
start "" /B "%REPO_ROOT%\.venv\Scripts\pythonw.exe" "%SCRIPT_DIR%screen_recorder.py"
exit
