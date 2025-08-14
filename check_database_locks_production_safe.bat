@echo off
echo ================================================================
echo PRODUCTION-SAFE MS ACCESS LOCK CHECKER
echo ================================================================
echo.
echo This script is SAFE to run in production environments.
echo It will NOT:
echo   - Kill running Access processes
echo   - Force-remove active database lock files  
echo   - Interfere with live databases
echo.
echo It WILL:
echo   - Check for potential lock issues
echo   - Safely clear COM cache
echo   - Provide production-safe recommendations
echo   - Test COM functionality without interference
echo.
echo ================================================================
echo.
pause

echo Running production-safe diagnostic...
python fix_database_locks_production_safe.py %1

echo.
echo ================================================================
echo DIAGNOSTIC COMPLETED
echo ================================================================
echo.
echo Next steps:
echo 1. Review the recommendations above
echo 2. If system is ready, run: run_enhanced_conversion.bat
echo 3. The converter has built-in retry mechanisms for lock issues
echo.
echo For production environments, consider:
echo - Running during off-peak hours
echo - Copying MDB files before conversion
echo - Using incremental approach (one database at a time)
echo.
pause
