@echo off
echo MS Access ODBC Driver Fix Script - OLD MDB FILES
echo ================================================
echo.

REM Check Python architecture
echo Checking Python architecture...
python -c "import platform; print('Python Architecture:', platform.architecture()[0])"
echo.

REM Check current ODBC drivers
echo Your current ODBC drivers show NO ACCESS DRIVERS:
echo   - QB
echo   - SQL Anywhere  
echo   - SQL Server
echo   - SQL Server Native Client 10.0
echo   - SQL Server Native Client 11.0
echo.

echo ❌ MISSING: Microsoft Access Driver (*.mdb, *.accdb)
echo.

echo SOLUTION FOR OLD .MDB FILES:
echo ===========================
echo.
echo Option 1: Microsoft Access Database Engine 2016 (RECOMMENDED)
echo -------------------------------------------------------------
echo 1. Download from: https://www.microsoft.com/en-us/download/details.aspx?id=54920
echo.
python -c "import platform; arch=platform.architecture()[0]; print('2. For your Python (' + arch + '):'); exe='AccessDatabaseEngine_X64.exe' if '64bit' in arch else 'AccessDatabaseEngine.exe'; print('   Download: ' + exe)"
echo.
echo 3. If installation fails with "Another version already installed":
python -c "import platform; arch=platform.architecture()[0]; exe='AccessDatabaseEngine_X64.exe' if '64bit' in arch else 'AccessDatabaseEngine.exe'; print('   Run as Administrator: ' + exe + ' /quiet')"
echo.

echo Option 2: Legacy Jet Database Engine 4.0 (For VERY old MDB)
echo -----------------------------------------------------------
echo 1. Download: https://www.microsoft.com/en-us/download/details.aspx?id=23734
echo 2. Install Microsoft Data Access Components (MDAC)
echo.

echo Option 3: Alternative Solutions
echo ------------------------------
echo A. Use the legacy converter: python legacy_mdb_converter.py
echo B. Install Microsoft Access (provides COM automation)
echo C. Use 32-bit Python + 32-bit drivers (better compatibility)
echo D. Manual export: Open in Access → Export each table to CSV
echo.

REM Offer to download automatically
set /p download="Open download page for Access Database Engine? (y/n): "
if /i "%download%"=="y" (
    start https://www.microsoft.com/en-us/download/details.aspx?id=54920
    echo Download page opened in your browser.
)
echo.

echo AFTER INSTALLATION:
echo ===================
echo 1. Close ALL command prompts and VS Code
echo 2. Open a NEW command prompt
echo 3. Test: python -c "import pyodbc; print([d for d in pyodbc.drivers() if 'access' in d.lower()])"
echo 4. Should show: ['Microsoft Access Driver (*.mdb, *.accdb)']
echo 5. Then run: python config_setup.py test
echo.

echo TROUBLESHOOTING:
echo ===============
echo - If still no drivers: Try installing both 32-bit and 64-bit versions
echo - For very old MDB: Use legacy_mdb_converter.py instead
echo - Check Windows Event Viewer for installation errors
echo - Try running installer as Administrator
echo.
pause
