@echo off
REM ============================================================================
REM NOTION ARTICLES SYNC AUTOMATION BATCH SCRIPT
REM ============================================================================
REM Description: This batch file runs the notion_articles_sync.py script
REM              which syncs articles database to archive with full sync capabilities.
REM 
REM Usage: Simply double-click this bat file or run it from command line.
REM ============================================================================

echo [INFO] Starting Notion Articles Sync process...

for %%I in ("%~dp0..\..") do set "REPO_ROOT=%%~fI"

REM Use the repo's own virtual environment (resolved relative to this file)
set "VENV_DIR=%REPO_ROOT%\.venv"

REM Set the path to the notion articles sync scripts
for %%I in ("%~dp0.") do set "SCRIPT_DIR=%%~fI"

set "VENV_PY=%VENV_DIR%\Scripts\python.exe"
if not exist "%VENV_PY%" (
    echo [ERROR] Virtual environment interpreter not found at "%VENV_PY%"
    pause
    exit /b 1
)

echo [INFO] Changing to script directory: "%SCRIPT_DIR%"
cd /d "%SCRIPT_DIR%"
if errorlevel 1 (
    echo [ERROR] Failed to change directory.
    exit /b 1
)

echo [INFO] Running notion_articles_sync.py...
"%VENV_PY%" notion_articles_sync.py
if errorlevel 1 (
    echo [ERROR] notion_articles_sync.py returned an error code.
    pause
    exit /b 1
)

echo [INFO] Process completed successfully!
pause

