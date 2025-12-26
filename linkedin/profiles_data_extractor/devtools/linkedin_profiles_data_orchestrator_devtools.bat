@echo off
REM LinkedIn Profiles Data Orchestrator - DevTools Version
REM
REM This script runs the DevTools-based profile data orchestrator.
REM Make sure Chrome is running with remote debugging enabled first!
REM
REM Usage: Run start_chrome_debug.bat first, then this script.

echo ========================================
echo LinkedIn Profiles Data Orchestrator
echo DevTools Version
echo ========================================
echo.

REM Check if Chrome is running with debugging
echo Checking if Chrome debug port is accessible...
powershell -Command "try { $response = Invoke-WebRequest -Uri 'http://localhost:9222/json' -TimeoutSec 5; if ($response.StatusCode -eq 200) { Write-Host '✅ Chrome debug port accessible' -ForegroundColor Green } } catch { Write-Host '❌ Chrome debug port not accessible' -ForegroundColor Red; Write-Host 'Please run start_chrome_debug.bat first!' -ForegroundColor Yellow; exit 1 }"

echo.
echo Starting DevTools-based profile orchestrator...
echo.

REM Get the directory where this batch file is located
set SCRIPT_DIR=%~dp0

REM Run the Python script
python "%SCRIPT_DIR%linkedin_profiles_data_orchestrator_devtools.py"

echo.
echo DevTools orchestration completed!
echo.
pause
