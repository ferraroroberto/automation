@echo off
REM LinkedIn Profile Data Extractor - Chrome Debug Launcher
REM
REM This batch file starts Google Chrome with remote debugging enabled,
REM allowing the DevTools-based profile checker to connect and extract data.
REM
REM Usage: start_chrome_debug.bat [profile_path]
REM
REM If profile_path is not provided, Chrome will use the default profile.

echo ========================================
echo LinkedIn Profile Data Extractor
echo Chrome Debug Launcher
echo ========================================
echo.

if "%~1"=="" (
    echo Starting Chrome with remote debugging on port 9222...
    echo Using default Chrome profile.
    echo.
    start "" "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --remote-allow-origins=* --user-data-dir="%USERPROFILE%\AppData\Local\Google\Chrome\User Data\DebugProfile"
) else (
    echo Starting Chrome with remote debugging on port 9222...
    echo Using custom profile: %~1
    echo.
    start "" "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --remote-allow-origins=* --user-data-dir="%~1"
)

echo.
echo Chrome started with remote debugging enabled!
echo.
echo Next steps:
echo 1. Open LinkedIn and navigate to your search results
echo 2. Open multiple tabs with different pages if needed
echo 3. Run the DevTools profile checker: linkedin_profile_search_checker_devtools.py
echo.
echo Press any key to exit...
pause >nul
