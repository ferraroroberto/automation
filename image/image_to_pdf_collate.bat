@echo off
chcp 65001 >nul
set "PROJECT_DIR=E:\automation\automation"
set "SCRIPT_PATH=%PROJECT_DIR%\image\image_to_pdf_collate.py"
set "VENV_PY=%PROJECT_DIR%\.venv\Scripts\python.exe"

cd /d "%PROJECT_DIR%"
if errorlevel 1 (echo [ERROR] Failed to change directory. & exit /b 1)

if not exist "%VENV_PY%" (echo [ERROR] Virtual environment interpreter not found at "%VENV_PY%". & exit /b 1)

"%VENV_PY%" "%SCRIPT_PATH%"
exit /b %ERRORLEVEL%
