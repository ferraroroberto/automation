@echo off
chcp 65001 >nul

REM ============================================================================
REM TRANSCRIBE VOICE SCRIPT LAUNCHER
REM ============================================================================
REM Description: This batch file launches the transcribe_voice.py Python script
REM              with the appropriate virtual environment.
REM 
REM Usage: Simply double-click this bat file or run it from command line.
REM        It will automatically activate the Python virtual environment
REM        before executing the script with the --launch_from_elgato parameter.
REM
REM Requirements:
REM   - Python installed on the system
REM   - Virtual environment at the specified location (.venv folder)
REM   - The transcribe_voice.py script file
REM
REM ============================================================================

set "PROJECT_DIR=E:\automation\automation"
set "SCRIPT_PATH=E:\automation\automation\audio\transcribe_voice_gui.py"
set "VENV_ACTIVATE=%PROJECT_DIR%\.venv\Scripts\activate.bat"

echo [INFO] Changing to project directory: "%PROJECT_DIR%\audio"
cd /d "E:\automation\automation"
if errorlevel 1 (
    echo [ERROR] Failed to change directory.
    exit /b 1
)

REM Check for virtual environment
if exist "%VENV_ACTIVATE%" (
    echo [INFO] Activating virtual environment: "%VENV_ACTIVATE%"
    call "%VENV_ACTIVATE%"
    if errorlevel 1 (
        echo [ERROR] Failed to activate virtual environment.
        exit /b 1
    )
) else (
    echo [INFO] No virtual environment found. Using system Python.
)

echo [INFO] Running script: "%SCRIPT_PATH%" with --launch_from_elgato
python "%SCRIPT_PATH%" --launch_from_elgato
set "PY_EXIT_CODE=%ERRORLEVEL%"

if not "%PY_EXIT_CODE%"=="0" (
    echo [ERROR] Python script exited with code %PY_EXIT_CODE%.
    echo.
    echo [INFO] Press any key to exit...
    pause >nul
) else (
    echo [INFO] Script executed successfully with exit code 0.
)

exit /b %PY_EXIT_CODE%