@echo off
pushd "%~dp0"
"%~dp0..\..\.venv\Scripts\python.exe" wifi_connect.py
popd
pause
