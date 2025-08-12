#!/usr/bin/env python3
"""
Diagnostic script for MS Access ODBC connection issues.
This script helps identify and resolve common ODBC driver problems.
"""

import sys
import platform
import subprocess
from pathlib import Path

try:
    import pyodbc
except ImportError:
    print("Error: pyodbc is not installed. Run: pip install pyodbc")
    sys.exit(1)


def check_python_architecture():
    """Check if Python is 32-bit or 64-bit."""
    arch = platform.architecture()[0]
    print(f"Python Architecture: {arch}")
    return arch


def list_odbc_drivers():
    """List all available ODBC drivers."""
    try:
        drivers = pyodbc.drivers()
        print(f"\nAvailable ODBC Drivers ({len(drivers)} total):")
        print("-" * 50)
        
        access_drivers = []
        other_drivers = []
        
        for driver in sorted(drivers):
            if 'access' in driver.lower() or 'mdb' in driver.lower() or 'accdb' in driver.lower():
                access_drivers.append(driver)
                print(f"✅ {driver} (Access-related)")
            else:
                other_drivers.append(driver)
        
        if not access_drivers:
            print("❌ No Microsoft Access drivers found!")
        
        print(f"\nOther drivers ({len(other_drivers)}):")
        for driver in other_drivers[:10]:  # Show first 10 to avoid clutter
            print(f"   {driver}")
        if len(other_drivers) > 10:
            print(f"   ... and {len(other_drivers) - 10} more")
        
        return access_drivers
        
    except Exception as e:
        print(f"Error listing drivers: {e}")
        return []


def test_access_connection(test_db_path=None):
    """Test connection to an Access database."""
    print(f"\nTesting Access Database Connection:")
    print("-" * 40)
    
    if not test_db_path:
        print("No test database specified - skipping connection test")
        return False
    
    test_path = Path(test_db_path)
    if not test_path.exists():
        print(f"Test database not found: {test_path}")
        return False
    
    access_drivers = [d for d in pyodbc.drivers() if 'access' in d.lower() or 'mdb' in d.lower()]
    
    if not access_drivers:
        print("❌ No Access drivers available for testing")
        return False
    
    for driver in access_drivers:
        try:
            conn_str = f"DRIVER={{{driver}}};DBQ={test_path.absolute()};ExtendedAnsiSQL=1;"
            print(f"Testing with driver: {driver}")
            
            conn = pyodbc.connect(conn_str, timeout=10)
            cursor = conn.cursor()
            
            # Try to list tables
            tables = []
            for table_info in cursor.tables(tableType='TABLE'):
                if not table_info.table_name.startswith('MSys'):
                    tables.append(table_info.table_name)
            
            conn.close()
            print(f"✅ Connection successful! Found {len(tables)} tables")
            if tables:
                print(f"   Sample tables: {', '.join(tables[:3])}")
            return True
            
        except Exception as e:
            print(f"❌ Connection failed with {driver}: {e}")
    
    return False


def check_access_engine_installation():
    """Check if Microsoft Access Database Engine is installed."""
    print(f"\nChecking Microsoft Access Database Engine:")
    print("-" * 50)
    
    # Check registry for installed components (Windows specific)
    if platform.system() == "Windows":
        try:
            import winreg
            
            # Check for Access Database Engine in registry
            registry_paths = [
                r"SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office",
                r"SOFTWARE\Microsoft\Office",
                r"SOFTWARE\WOW6432Node\Microsoft\Office"
            ]
            
            found_versions = []
            for base_path in registry_paths:
                try:
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, base_path) as key:
                        i = 0
                        while True:
                            try:
                                subkey_name = winreg.EnumKey(key, i)
                                if subkey_name.replace(".", "").isdigit():  # Version numbers like 16.0, 15.0
                                    try:
                                        subkey_path = f"{base_path}\\{subkey_name}\\Access Connectivity Engine"
                                        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, subkey_path):
                                            found_versions.append(f"Access Database Engine {subkey_name}")
                                    except FileNotFoundError:
                                        pass
                                i += 1
                            except OSError:
                                break
                except FileNotFoundError:
                    continue
            
            if found_versions:
                print("✅ Found Microsoft Access Database Engine installations:")
                for version in found_versions:
                    print(f"   - {version}")
            else:
                print("❌ Microsoft Access Database Engine not found in registry")
                
        except ImportError:
            print("Cannot check registry (winreg not available)")
    else:
        print("Registry check only available on Windows")


def provide_solutions():
    """Provide solutions for common issues."""
    python_arch = check_python_architecture()
    
    print(f"\n" + "=" * 60)
    print("SOLUTIONS FOR ODBC CONNECTION ISSUES")
    print("=" * 60)
    
    print(f"\n1. Install Microsoft Access Database Engine")
    print("-" * 45)
    print("Download from Microsoft:")
    print("https://www.microsoft.com/en-us/download/details.aspx?id=54920")
    print()
    
    if "64bit" in python_arch:
        print("⚠️  You're using 64-bit Python")
        print("   - Download and install the 64-bit version (AccessDatabaseEngine_X64.exe)")
        print("   - If you get 'Another version is installed' error:")
        print("     Run: AccessDatabaseEngine_X64.exe /quiet")
    else:
        print("⚠️  You're using 32-bit Python")
        print("   - Download and install the 32-bit version (AccessDatabaseEngine.exe)")
        print("   - If you get 'Another version is installed' error:")
        print("     Run: AccessDatabaseEngine.exe /quiet")
    
    print(f"\n2. Alternative: Use 32-bit Python with 32-bit Access Engine")
    print("-" * 55)
    print("If you continue having issues:")
    print("   - Install 32-bit Python from python.org")
    print("   - Install 32-bit Access Database Engine")
    print("   - 32-bit combinations are generally more compatible")
    
    print(f"\n3. Verify Installation")
    print("-" * 25)
    print("After installing the Access Database Engine:")
    print("   - Restart your command prompt/IDE")
    print("   - Run this diagnostic script again")
    print("   - Test with a sample Access database")
    
    print(f"\n4. Alternative Solutions")
    print("-" * 25)
    print("If ODBC continues to fail:")
    print("   - Consider using mdb-tools (Linux/Mac)")
    print("   - Use Access automation via COM (Windows only)")
    print("   - Convert databases using Access itself")
    print("   - Use third-party tools like MDB Viewer Plus")


def main():
    """Main diagnostic function."""
    print("=" * 60)
    print("MS ACCESS ODBC DIAGNOSTIC TOOL")
    print("=" * 60)
    
    # Basic system info
    print(f"Operating System: {platform.system()} {platform.release()}")
    check_python_architecture()
    
    # Check ODBC drivers
    access_drivers = list_odbc_drivers()
    
    # Check Access Engine installation
    check_access_engine_installation()
    
    # Test connection if user provides a database
    test_db = input(f"\nEnter path to test Access database (optional, press Enter to skip): ").strip()
    if test_db:
        test_access_connection(test_db)
    
    # Provide solutions
    provide_solutions()
    
    print(f"\n" + "=" * 60)
    print("SUMMARY")
    print("=" * 60)
    
    if access_drivers:
        print(f"✅ Found {len(access_drivers)} Access driver(s)")
        print("Your system should be able to connect to Access databases")
    else:
        print("❌ No Access drivers found")
        print("You need to install Microsoft Access Database Engine")
    
    print(f"\nFor more help, check the logs when running the converter")
    print("or visit: https://github.com/mkleehammer/pyodbc/wiki")


if __name__ == "__main__":
    main()
    input("\nPress Enter to exit...")
