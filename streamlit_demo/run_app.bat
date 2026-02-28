@echo off
REM ==========================================================
REM  Streamlit Demo Playground – Windows Launcher
REM ==========================================================
REM  Double-click this file to start the application.
REM  It activates a virtual environment (.venv) and launches
REM  the Streamlit app.
REM  and launches the Streamlit app.
REM ==========================================================

title Streamlit Demo Playground

REM -- Resolve the directory where this .bat file lives ------
cd /d "%~dp0"

REM -- Locate virtual environment -----------------------------
set "VENV_DIR="
if exist ".venv\Scripts\activate.bat" set "VENV_DIR=%cd%\.venv"
if not defined VENV_DIR if exist "..\.venv\Scripts\activate.bat" set "VENV_DIR=%cd%\..\.venv"

REM -- Check that .venv exists --------------------------------
if not defined VENV_DIR (
    echo.
    echo  [ERROR] Virtual environment not found.
    echo  Please create it first:
    echo.
    echo      cd /d "%cd%\.."
    echo      python -m venv .venv
    echo      .venv\Scripts\activate
    echo      pip install -r requirements.txt
    echo.
    pause
    exit /b 1
)

REM -- Activate the virtual environment -----------------------
call "%VENV_DIR%\Scripts\activate.bat"

REM -- Launch Streamlit ---------------------------------------
echo.
echo  Starting Streamlit Demo Playground...
echo  Press Ctrl+C to stop the server.
echo.
streamlit run main_menu.py

REM -- Keep the window open if Streamlit exits ----------------
pause
