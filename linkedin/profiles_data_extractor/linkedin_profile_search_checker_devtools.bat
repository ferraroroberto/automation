@echo off
REM LinkedIn Profile Search Checker - DevTools Version
REM
REM This script runs the Chrome DevTools-based profile checker.
REM Make sure Chrome is running with remote debugging enabled first!
REM
REM Usage: Run start_chrome_debug.bat first, then this script.

echo ========================================
echo LinkedIn Profile Search Checker
echo Chrome DevTools Version
echo ========================================
echo.

REM Check if Chrome is running with debugging
echo Checking if Chrome debug port is accessible...
powershell -Command "try { $response = Invoke-WebRequest -Uri 'http://localhost:9222/json' -TimeoutSec 5; if ($response.StatusCode -eq 200) { Write-Host '✅ Chrome debug port accessible' -ForegroundColor Green } } catch { Write-Host '❌ Chrome debug port not accessible' -ForegroundColor Red; Write-Host 'Please run start_chrome_debug.bat first!' -ForegroundColor Yellow; exit 1 }"

echo.
echo Starting DevTools-based profile checker...
echo.

REM Get the directory where this batch file is located
set SCRIPT_DIR=%~dp0

REM Run the Python script
python "%SCRIPT_DIR%linkedin_profile_search_checker_devtools.py"

echo.
echo DevTools profile check completed!
echo.
pause
