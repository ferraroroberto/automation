@echo off
rem One poll cycle (heartbeat target for the scheduler). Runs from the repo root with the repo venv.
cd /d "%~dp0.."
set PYTHONUTF8=1
".venv\Scripts\python.exe" -m parking.poller %*
exit /b %ERRORLEVEL%
