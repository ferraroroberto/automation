@echo off
REM Activates the virtual environment and runs keycaster.py

REM Path to your virtual environment
set VENV_DIR=E:\onedrive\Documentos\Roberto\projects\automation\notion-automation\local\.venv_keycaster

REM Activate the venv
call %VENV_DIR%\Scripts\activate.bat

REM Run the script with Python from the venv
python keycaster.py

REM Pause so you can read any error messages
pause
