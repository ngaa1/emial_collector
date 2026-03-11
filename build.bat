@echo off

echo Start building Outlook email reader tool...
echo ================================

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Error: Python not found. Please install Python 3.6+ first.
    pause
    exit /b 1
)

REM Check if pip is available
pip --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Error: pip not found. Please ensure Python is installed correctly.
    pause
    exit /b 1
)

REM Install dependencies
echo Installing dependencies...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo Error: Failed to install dependencies.
    pause
    exit /b 1
)

REM Install PyInstaller
echo Installing PyInstaller...
pip install pyinstaller
if %errorlevel% neq 0 (
    echo Error: Failed to install PyInstaller.
    pause
    exit /b 1
)

REM Build application
echo Building application...
pyinstaller --onefile outlook_reader.py
if %errorlevel% neq 0 (
    echo Error: Build failed.
    pause
    exit /b 1
)

echo ================================
echo Build completed!
echo Executable file has been generated in the dist directory
echo You can copy dist\outlook_reader.exe to other computers for use
echo Note: Please ensure Outlook is installed and running on the target computer before using
echo ================================
pause
