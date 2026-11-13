@echo off
REM Launch Snap (hotkey region screenshot) into the system tray using the project venv, no console window.
set "SCRIPT_DIR=%~dp0"
set "REPO_ROOT=%SCRIPT_DIR%..\.."
if not exist "%SCRIPT_DIR%config.json" copy "%SCRIPT_DIR%config.sample.json" "%SCRIPT_DIR%config.json" >nul
start "" /B "%REPO_ROOT%\.venv\Scripts\pythonw.exe" "%SCRIPT_DIR%snap_tool.py"
exit
