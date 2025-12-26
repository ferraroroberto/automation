@echo off
REM ============================================================================
REM LINKEDIN PROFILES DATA ORCHESTRATOR - DEVTOOLS VERSION
REM ============================================================================
REM Description: This batch file runs the DevTools-based profile data orchestrator.
REM Make sure Chrome is running with remote debugging enabled first!
REM
REM Usage: Run start_chrome_debug.bat first, then this script.
REM ============================================================================

echo ========================================
echo LinkedIn Profiles Data Orchestrator
echo DevTools Version
echo ========================================
echo.

REM Check if Chrome is running with debugging
echo Checking if Chrome debug port is accessible...
powershell -Command "try { $response = Invoke-WebRequest -Uri 'http://localhost:9222/json' -TimeoutSec 5; if ($response.StatusCode -eq 200) { Write-Host '[OK] Chrome debug port accessible' -ForegroundColor Green } } catch { Write-Host '[ERROR] Chrome debug port not accessible' -ForegroundColor Red; Write-Host 'Please run start_chrome_debug.bat first!' -ForegroundColor Yellow; exit 1 }"

echo.
echo Starting DevTools-based profile orchestrator...
echo.

REM Set the path to the virtual environment
set "VENV_DIR=E:\automation\automation\.venv"

REM Get the directory where this batch file is located
set "SCRIPT_DIR=%~dp0"

echo [INFO] Activating virtual environment...
call "%VENV_DIR%\Scripts\activate.bat"
if errorlevel 1 (
    echo [ERROR] Failed to activate virtual environment. Make sure it exists at %VENV_DIR%
    echo [INFO] Attempting to continue without virtual environment activation...
)

echo [INFO] Changing to script directory: "%SCRIPT_DIR%"
cd /d "%SCRIPT_DIR%"
if errorlevel 1 (
    echo [ERROR] Failed to change directory.
    exit /b 1
)

REM Run the Python script
echo [INFO] Running linkedin_profiles_data_orchestrator_devtools.py...
python "%SCRIPT_DIR%linkedin_profiles_data_orchestrator_devtools.py"

if errorlevel 1 (
    echo [ERROR] Script failed with error code %errorlevel%
    echo [INFO] Press any key to see error details...
    pause >nul
    goto :eof
)

echo.
echo [INFO] DevTools orchestration completed!
echo.
pause
