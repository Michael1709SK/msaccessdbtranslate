@echo off
echo MS Access to MySQL Database Converter
echo =====================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Error: Python is not installed or not in PATH
    echo Please install Python 3.7 or higher from https://python.org
    pause
    exit /b 1
)

echo Checking Python installation...
python --version

REM Check if pip is available
pip --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Error: pip is not available
    pause
    exit /b 1
)

echo Installing required packages...
pip install -r requirements.txt

if %errorlevel% neq 0 (
    echo Error: Failed to install required packages
    pause
    exit /b 1
)

echo.
echo Installation completed successfully!
echo.
echo Next steps:
echo 1. Run: python config_setup.py setup
echo 2. Then: python run_converter.py
echo.
pause
