@echo off
REM ==========================================================
REM  Streamlit Demo Playground – Windows Launcher
REM ==========================================================
REM  Double-click this file to start the application.
REM  It activates the virtual environment located in .venv\
REM  and launches the Streamlit app.
REM ==========================================================

title Streamlit Demo Playground

REM -- Resolve the directory where this .bat file lives ------
cd /d "%~dp0"

REM -- Check that .venv exists --------------------------------
if not exist ".venv\Scripts\activate.bat" (
    echo.
    echo  [ERROR] Virtual environment not found at .venv\
    echo  Please create it first:
    echo.
    echo      python -m venv .venv
    echo      .venv\Scripts\activate
    echo      pip install -r requirements.txt
    echo.
    pause
    exit /b 1
)

REM -- Activate the virtual environment -----------------------
call .venv\Scripts\activate.bat

REM -- Launch Streamlit ---------------------------------------
echo.
echo  Starting Streamlit Demo Playground...
echo  Press Ctrl+C to stop the server.
echo.
streamlit run app.py

REM -- Keep the window open if Streamlit exits ----------------
pause
