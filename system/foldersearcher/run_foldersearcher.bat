@echo off
REM Activates the virtual environment and runs foldersearcher.py

REM Path to your virtual environment
set VENV_DIR=E:\automation\automation\.venv\

REM Activate the venv
call %VENV_DIR%\Scripts\activate.bat

REM Run the script with Python from the venv
python foldersearcher.py

REM Pause so you can read any error messages
pause 