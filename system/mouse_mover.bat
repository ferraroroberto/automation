@echo off
REM Launch Mouse Mover silently into the system tray using the project venv.
set "SCRIPT_DIR=%~dp0"
set "REPO_ROOT=%SCRIPT_DIR%.."
start "" /B "%REPO_ROOT%\.venv\Scripts\pythonw.exe" "%SCRIPT_DIR%mouse_mover.py" --up 100 --right 1 --down 100 --left 1 --speed 0.001
exit
