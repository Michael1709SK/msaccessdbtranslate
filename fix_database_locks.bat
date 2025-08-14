@echo off
echo ================================================================
echo MS ACCESS DATABASE LOCK FIXER
echo ================================================================
echo.
echo This will fix "database already open" errors by:
echo - Closing any running Access processes
echo - Removing lock files (*.ldb, *.laccdb)
echo - Clearing COM automation cache
echo - Testing Access COM functionality
echo.
echo Make sure to close any Access applications before continuing.
echo.
pause

python fix_database_locks.py

echo.
echo ================================================================
echo CLEANUP COMPLETED
echo ================================================================
echo.
echo You can now run the conversion again:
echo   run_enhanced_conversion.bat
echo.
echo Or directly:
echo   python access_com_converter.py "C:\your\mdb\path" --user username --password password
echo.
pause
