@echo off
REM ============================================================================
REM NOTION JOURNAL AUTOMATION BATCH SCRIPT
REM ============================================================================
REM Description: This batch file runs three Python scripts in sequence:
REM              1. notion_databases_dump.py - Downloads databases from Notion
REM              2. notion_databases_clean.py - Cleans the downloaded databases
REM              3. journal_automation.py - Processes journal entries
REM              Then opens the journal output folder in Explorer.
REM 
REM Usage: Simply double-click this bat file or run it from command line.
REM ============================================================================

echo [INFO] Starting Notion Journal Automation process...

REM Set the path to the virtual environment
set "VENV_DIR=E:\automation\automation\.venv"

REM Set the path to the notion scripts
set "SCRIPT_DIR=E:\automation\automation\notion"

REM Set the path to the journal output folder
set "JOURNAL_OUTPUT=E:\automation\notion-automation-files\journal-output"

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

echo [INFO] Running notion_databases_dump.py...
python notion_databases_dump.py
if errorlevel 1 (
    echo [WARNING] notion_databases_dump.py returned an error code. Continuing anyway...
)

echo [INFO] Running notion_databases_clean.py...
python notion_databases_clean.py
if errorlevel 1 (
    echo [WARNING] notion_databases_clean.py returned an error code. Continuing anyway...
)

echo [INFO] Running journal_automation.py...
python journal_automation.py
if errorlevel 1 (
    echo [WARNING] journal_automation.py returned an error code. Continuing anyway...
)

echo [INFO] Opening journal output folder in Explorer...
explorer "%JOURNAL_OUTPUT%"

echo [INFO] Process completed successfully!
pause
