@echo off
setlocal enabledelayedexpansion

REM Minimal build script for PDF to Images Converter (Windows)
REM This script creates a temporary virtual environment, installs only the required dependencies,
REM builds the executable, and cleans up everything except the final .exe.

echo Building minimal PDF to Images Converter executable...
echo.

REM Get the project directory
set "PROJECT_DIR=%~dp0"
set "PROJECT_DIR=%PROJECT_DIR:~0,-1%"
set "TEMP_VENV=%PROJECT_DIR%\.build_venv"
set "DIST_DIR=%PROJECT_DIR%\dist"
set "MAIN_SCRIPT=%PROJECT_DIR%\pdf_to_images_converter.py"

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo Error: Python is not installed or not in PATH
    pause
    exit /b 1
)

REM 1. Create temporary virtual environment
echo Creating temporary virtual environment...
if exist "%TEMP_VENV%" (
    echo Removing existing temporary environment...
    rmdir /s /q "%TEMP_VENV%" 2>nul
)
python -m venv "%TEMP_VENV%"
if errorlevel 1 (
    echo Error: Failed to create virtual environment
    pause
    exit /b 1
)

REM 2. Activate virtual environment
echo Activating virtual environment...
call "%TEMP_VENV%\Scripts\activate.bat"
if errorlevel 1 (
    echo Error: Failed to activate virtual environment
    pause
    exit /b 1
)

REM 3. Upgrade pip and install only required dependencies
echo Upgrading pip...
python -m pip install --upgrade pip
if errorlevel 1 (
    echo Error: Failed to upgrade pip
    pause
    exit /b 1
)

echo Installing required dependencies...
pip install PyMuPDF==1.23.8 pyinstaller==6.3.0
if errorlevel 1 (
    echo Error: Failed to install dependencies
    pause
    exit /b 1
)

REM 4. Build the executable using PyInstaller with console support
echo Building executable...
echo This may take a few minutes...
pyinstaller --onefile --console --name "PDF_to_Images_Converter_Optimized" --clean --noconfirm --log-level=WARN "%MAIN_SCRIPT%"
if errorlevel 1 (
    echo Error: Failed to build executable
    echo.
    echo Common solutions:
    echo 1. Make sure you have enough disk space
    echo 2. Try running as administrator
    echo 3. Check if antivirus is blocking the build
    pause
    exit /b 1
)

REM 5. Copy the resulting .exe to the dist folder (if not already there)
if not exist "%DIST_DIR%" (
    mkdir "%DIST_DIR%"
)
if exist "%PROJECT_DIR%\dist\PDF_to_Images_Converter_Optimized.exe" (
    copy /Y "%PROJECT_DIR%\dist\PDF_to_Images_Converter_Optimized.exe" "%DIST_DIR%\" >nul
    echo Executable copied to: %DIST_DIR%\PDF_to_Images_Converter_Optimized.exe
) else (
    echo Warning: Executable not found in expected location
)

REM 6. Deactivate and remove the temporary virtual environment and build artifacts
echo Cleaning up...
call "%TEMP_VENV%\Scripts\deactivate.bat" 2>nul
rmdir /s /q "%TEMP_VENV%" 2>nul
rmdir /s /q "%PROJECT_DIR%\build" 2>nul
rmdir /s /q "%PROJECT_DIR%\__pycache__" 2>nul
del "%PROJECT_DIR%\*.spec" 2>nul

REM 7. Done
echo.
echo Minimal build completed! 
if exist "%DIST_DIR%\PDF_to_Images_Converter_Optimized.exe" (
    echo Executable is in: %DIST_DIR%\PDF_to_Images_Converter_Optimized.exe
    echo Size: 
    for %%A in ("%DIST_DIR%\PDF_to_Images_Converter_Optimized.exe") do echo   %%~zA bytes
) else (
    echo Warning: Executable not found in dist directory
)
echo.
pause
