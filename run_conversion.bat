@echo off
echo MS Access to MySQL Database Converter - Quick Start
echo ==================================================
echo.

REM Check if config file exists
if not exist "converter_config.json" (
    echo No configuration found. Running setup...
    python config_setup.py setup
    if %errorlevel% neq 0 (
        echo Setup failed. Please check the errors above.
        pause
        exit /b 1
    )
)

echo.
echo Starting conversion with saved configuration...
python run_converter.py

if %errorlevel% equ 0 (
    echo.
    echo Conversion completed successfully!
) else (
    echo.
    echo Conversion completed with errors. Check the log files for details.
)

echo.
pause
