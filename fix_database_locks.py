#!/usr/bin/env python3
"""
Quick fix for 'database already open' errors in MS Access COM automation.
Run this if you encounter "database already open" errors during conversion.
"""

import os
import sys
import time
import subprocess

def kill_access_processes():
    """Kill any running Microsoft Access processes."""
    print("🔍 Checking for running Microsoft Access processes...")
    
    try:
        # Kill MSACCESS.EXE processes
        result = subprocess.run(['taskkill', '/F', '/IM', 'MSACCESS.EXE'], 
                              capture_output=True, text=True)
        if result.returncode == 0:
            print("✅ Killed Microsoft Access processes")
        else:
            print("ℹ️  No Microsoft Access processes found")
    except Exception as e:
        print(f"⚠️  Could not kill Access processes: {e}")

def clear_access_locks():
    """Clear potential Access lock files."""
    print("🔍 Checking for Access lock files...")
    
    # Common lock file patterns
    lock_patterns = ['*.ldb', '*.laccdb', '~$*.mdb', '~$*.accdb']
    
    current_dir = os.getcwd()
    found_locks = False
    
    for pattern in lock_patterns:
        try:
            import glob
            lock_files = glob.glob(os.path.join(current_dir, '**', pattern), recursive=True)
            
            for lock_file in lock_files:
                try:
                    os.remove(lock_file)
                    print(f"🗑️  Removed lock file: {os.path.basename(lock_file)}")
                    found_locks = True
                except Exception as e:
                    print(f"⚠️  Could not remove {os.path.basename(lock_file)}: {e}")
        except Exception:
            pass
    
    if not found_locks:
        print("ℹ️  No lock files found")

def clear_com_cache():
    """Clear COM automation cache."""
    print("🧹 Clearing COM cache...")
    
    try:
        import win32com.client
        # Clear the COM cache
        win32com.client.gencache.GetGeneratePath()
        import shutil
        cache_dir = win32com.client.gencache.GetGeneratePath()
        
        if os.path.exists(cache_dir):
            shutil.rmtree(cache_dir, ignore_errors=True)
            print("✅ COM cache cleared")
        else:
            print("ℹ️  COM cache directory not found")
    except Exception as e:
        print(f"⚠️  Could not clear COM cache: {e}")

def test_access_com():
    """Test if Access COM automation is working."""
    print("🧪 Testing Access COM automation...")
    
    try:
        import win32com.client
        
        # Try to create Access application
        access_app = win32com.client.Dispatch('Access.Application')
        print("✅ Access COM object created successfully")
        
        # Try to quit cleanly
        access_app.Quit()
        access_app = None
        print("✅ Access COM object closed successfully")
        
        return True
    except Exception as e:
        print(f"❌ Access COM test failed: {e}")
        return False

def main():
    """Main function to fix Access database issues."""
    print("🔧 MS ACCESS DATABASE LOCK FIXER")
    print("=" * 50)
    print("This script will fix common 'database already open' errors")
    print("=" * 50)
    
    # Step 1: Kill any running Access processes
    kill_access_processes()
    
    # Step 2: Clear lock files
    clear_access_locks()
    
    # Step 3: Clear COM cache
    clear_com_cache()
    
    # Step 4: Wait a moment
    print("⏳ Waiting for cleanup to complete...")
    time.sleep(3)
    
    # Step 5: Test Access COM
    if test_access_com():
        print("\n✅ ACCESS COM AUTOMATION FIXED!")
        print("You can now run the conversion again:")
        print("python access_com_converter.py [your_parameters]")
    else:
        print("\n❌ Issues remain. Try these additional steps:")
        print("1. Restart your computer")
        print("2. Run this script as Administrator")
        print("3. Check if Access is properly installed")
        print("4. Try running: regsvr32 msaccess.exe")
    
    print("\n🔄 If problems persist, the enhanced converter includes:")
    print("- Automatic database closing between operations")
    print("- Delays to allow proper cleanup") 
    print("- Safe database opening/closing methods")
    
    input("\nPress Enter to exit...")

if __name__ == "__main__":
    main()
