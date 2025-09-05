@echo off
echo CUDA Test Script Runner
echo =======================
echo.

cd /d E:\automation\automation

echo Running CUDA test with virtual environment...
echo.

call .\venv\Scripts\python.exe audio\transcribe_voice\test_cuda.py

echo.
echo Test completed. Press any key to exit...
pause > nul
