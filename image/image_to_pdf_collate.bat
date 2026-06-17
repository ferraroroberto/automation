@echo off
chcp 65001 >nul
set "PROJECT_DIR=E:\automation\automation"
set "SCRIPT_PATH=%PROJECT_DIR%\image\image_to_pdf_collate.py"
set "VENV_ACTIVATE=%PROJECT_DIR%\.venv\Scripts\activate.bat"

cd /d "%PROJECT_DIR%"
if errorlevel 1 (echo [ERROR] Failed to change directory. & exit /b 1)

if exist "%VENV_ACTIVATE%" (call "%VENV_ACTIVATE%")

python "%SCRIPT_PATH%"
exit /b %ERRORLEVEL%
