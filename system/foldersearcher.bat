@echo off
REM Launch Folder Searcher silently into the system tray using the project venv.
set "SCRIPT_DIR=%~dp0"
set "REPO_ROOT=%SCRIPT_DIR%.."
start "" /B "%REPO_ROOT%\.venv\Scripts\pythonw.exe" "%SCRIPT_DIR%foldersearcher\foldersearcher.py"
exit
