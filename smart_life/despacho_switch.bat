@echo off
REM Toggles "luz despacho": if ON turns OFF, if OFF turns ON.
REM Uses the project virtual environment (same pattern as run_foldersearcher.bat).

set VENV_DIR=E:\automation\automation\.venv\

call %VENV_DIR%\Scripts\activate.bat

python light_control.py switch --light "luz despacho"

pause
