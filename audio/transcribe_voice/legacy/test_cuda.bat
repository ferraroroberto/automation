@echo off
echo CUDA Test Script Runner
echo =======================
echo.

cd /d E:\automation\automation

echo Running CUDA test with virtual environment...
echo.

REM Read VENV_FOLDER from .env file
for /f "tokens=2 delims==" %%i in ('findstr "VENV_FOLDER" .env') do set VENV_FOLDER=%%i

REM Run the CUDA test using the virtual environment
call "%VENV_FOLDER%\Scripts\python.exe" audio\transcribe_voice\test_cuda.py

echo.
echo Test completed. Press any key to exit...
pause > nul
