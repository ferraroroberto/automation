@echo off
REM Launch Markdown Preview silently into the system tray using the project venv.
REM Optional first argument: a .md file to open on launch.
set "SCRIPT_DIR=%~dp0"
set "REPO_ROOT=%SCRIPT_DIR%.."
start "" /B "%REPO_ROOT%\.venv\Scripts\pythonw.exe" "%SCRIPT_DIR%markdown_preview.py" %*
exit
