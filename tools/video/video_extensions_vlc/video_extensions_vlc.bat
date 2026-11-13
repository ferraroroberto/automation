@echo off
REM Video Extensions VLC Default Setter Launcher
REM This batch file launches the PowerShell script to set VLC as default for video extensions

REM Check if this is the elevated instance
if "%1"=="elevated" goto :run_script

echo Video Extensions VLC Default Setter
echo ===================================
echo.

REM Check if running as administrator
net session >nul 2>&1
if %errorLevel% == 0 (
    echo Running with administrator privileges - good!
    goto :run_script
) else (
    echo Warning: Not running as administrator. Some file associations require admin rights.
    echo.
    echo Attempting to restart with administrator privileges...
    echo.

    REM Restart as administrator with elevated flag and exit
    powershell "Start-Process '%~f0' -ArgumentList 'elevated' -Verb RunAs -Wait"
    exit /b
)

:run_script

:run_script
REM This is running in the elevated instance
if "%1"=="elevated" (
    echo Video Extensions VLC Default Setter
    echo ===================================
    echo.
    echo Running with administrator privileges - good!
)

REM Get the directory where this batch file is located
set "SCRIPT_DIR=%~dp0"
set "PS_SCRIPT=%SCRIPT_DIR%video_extensions_vlc.ps1"

REM Check if PowerShell script exists
if not exist "%PS_SCRIPT%" (
    echo Error: PowerShell script not found at %PS_SCRIPT%
    echo Make sure video_extensions_vlc.ps1 is in the same directory as this batch file.
    pause
    exit /b 1
)

echo Starting video extensions check...
echo.

REM Run the PowerShell script
powershell.exe -ExecutionPolicy Bypass -File "%PS_SCRIPT%"

echo.
echo Script completed.
pause