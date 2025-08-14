@echo off
echo =====================================================
echo MS Access to MySQL Enhanced Converter
echo =====================================================
echo.

:: Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python 3.7+ and try again
    pause
    exit /b 1
)

:: Check if required packages are installed
echo Checking dependencies...
python -c "import win32com.client, pandas, mysql.connector, tqdm" >nul 2>&1
if errorlevel 1 (
    echo Installing required packages...
    pip install pywin32 pandas mysql-connector-python tqdm
    if errorlevel 1 (
        echo ERROR: Failed to install dependencies
        pause
        exit /b 1
    )
)

:: Get user input
set /p SOURCE_DIR="Enter the path to your MDB files directory: "
if not exist "%SOURCE_DIR%" (
    echo ERROR: Directory "%SOURCE_DIR%" does not exist
    pause
    exit /b 1
)

set /p MYSQL_HOST="Enter MySQL host (default: localhost): "
if "%MYSQL_HOST%"=="" set MYSQL_HOST=localhost

set /p MYSQL_USER="Enter MySQL username: "
if "%MYSQL_USER%"=="" (
    echo ERROR: MySQL username is required
    pause
    exit /b 1
)

set /p MYSQL_PASSWORD="Enter MySQL password: "
if "%MYSQL_PASSWORD%"=="" (
    echo ERROR: MySQL password is required
    pause
    exit /b 1
)

echo.
echo =====================================================
echo Starting Enhanced Conversion Process
echo =====================================================
echo Source Directory: %SOURCE_DIR%
echo MySQL Host: %MYSQL_HOST%
echo MySQL User: %MYSQL_USER%
echo.
echo Features:
echo - Automatic table size detection and prioritization
echo - Progress bars and statistics tracking  
echo - Update existing tables with new data
echo - Comprehensive logging and reporting
echo - Graceful handling of large tables (1M+ rows)
echo.
echo Press Ctrl+C at any time to stop and generate final report
echo =====================================================
echo.

:: Create logs directory
if not exist "logs" mkdir logs

:: Run the enhanced converter
python access_com_converter.py "%SOURCE_DIR%" --host "%MYSQL_HOST%" --user "%MYSQL_USER%" --password "%MYSQL_PASSWORD%" --log-dir logs --update-interval 15

:: Show completion message
echo.
echo =====================================================
echo Conversion Process Completed
echo =====================================================
echo.
echo Check the following files for results:
echo - conversion_report_*.json (detailed JSON report)
echo - conversion_summary_*.txt (human-readable summary)
echo - logs\conversion_stats_*.log (detailed statistics log)
echo - logs\*.log (individual database conversion logs)
echo.
pause
