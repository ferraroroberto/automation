@echo off
REM Text Expander Launcher Script
REM Launches the text expander application using the local virtual environment

echo Starting Text Expander Application...

REM Run from the repo root so `-m system.textexpander.main` resolves the package.
REM The modules use intra-package relative imports, so they must be imported as
REM part of the `system.textexpander` package -- `python main.py` cannot work.
cd /d "%~dp0..\.."

REM Check if virtual environment exists
if not exist ".venv\Scripts\python.exe" (
    echo Virtual environment not found at %CD%\.venv\Scripts\python.exe
    echo Please ensure the virtual environment is set up correctly
    echo Run: py -m venv .venv
    echo Then install dependencies: .venv\Scripts\python.exe -m pip install -r system\textexpander\requirements_text_expander.txt
    pause
    exit /b 1
)

REM Check if requirements are installed
echo Checking dependencies...
.venv\Scripts\python.exe -c "import pystray, pynput, pyperclip" 2>nul
if errorlevel 1 (
    echo Some dependencies may be missing. Installing...
    .venv\Scripts\python.exe -m pip install -r system\textexpander\requirements_text_expander.txt
    if errorlevel 1 (
        echo Failed to install dependencies
        pause
        exit /b 1
    )
)

REM Launch the application
echo Launching Text Expander...
.venv\Scripts\python.exe -m system.textexpander.main

REM Check exit code
if %errorlevel% equ 0 (
    echo Text Expander closed successfully
) else (
    echo Text Expander exited with error code %errorlevel%
)

echo Text Expander session ended
pause
