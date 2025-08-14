#!/usr/bin/env python3
"""
PRODUCTION-SAFE fix for 'database already open' errors in MS Access COM automation.
This version is safe to run in production environments with live databases.
"""

import os
import sys
import time

def check_access_processes():
    """Check for running Microsoft Access processes WITHOUT killing them."""
    print("🔍 Checking for running Microsoft Access processes...")
    
    try:
        import subprocess
        result = subprocess.run(['tasklist', '/FI', 'IMAGENAME eq MSACCESS.EXE'], 
                              capture_output=True, text=True)
        
        if 'MSACCESS.EXE' in result.stdout:
            print("⚠️  WARNING: Microsoft Access is currently running!")
            print("   This may cause 'database already open' errors.")
            print("   Consider closing Access applications if they're not needed.")
            return True
        else:
            print("✅ No Microsoft Access processes detected")
            return False
    except Exception as e:
        print(f"ℹ️  Could not check Access processes: {e}")
        return False

def check_lock_files_in_source_only(source_dir=None):
    """Check for Access lock files ONLY in the source directory being processed."""
    print("🔍 Checking for Access lock files in source directory...")
    
    if not source_dir:
        source_dir = input("Enter the source directory path for MDB files: ").strip('"')
    
    if not os.path.exists(source_dir):
        print(f"❌ Source directory not found: {source_dir}")
        return False
    
    # Lock file patterns - only check in source directory
    lock_patterns = ['*.ldb', '*.laccdb']
    
    found_locks = []
    
    for pattern in lock_patterns:
        try:
            import glob
            # Only search in the specific source directory, not recursively
            lock_files = glob.glob(os.path.join(source_dir, pattern))
            found_locks.extend(lock_files)
        except Exception:
            pass
    
    if found_locks:
        print(f"⚠️  Found {len(found_locks)} lock files in source directory:")
        for lock_file in found_locks:
            file_size = os.path.getsize(lock_file) if os.path.exists(lock_file) else 0
            modified_time = time.ctime(os.path.getmtime(lock_file)) if os.path.exists(lock_file) else "Unknown"
            print(f"   📄 {os.path.basename(lock_file)} ({file_size} bytes, modified: {modified_time})")
        
        print("\n💡 PRODUCTION-SAFE OPTIONS:")
        print("1. These may be old lock files from previous sessions")
        print("2. If databases are NOT currently open, it's safe to remove them")
        print("3. If databases ARE currently open, DO NOT remove lock files")
        
        user_choice = input("\nAre you sure these databases are NOT currently open? (y/N): ").lower()
        
        if user_choice == 'y':
            removed_count = 0
            for lock_file in found_locks:
                try:
                    os.remove(lock_file)
                    print(f"🗑️  Removed: {os.path.basename(lock_file)}")
                    removed_count += 1
                except Exception as e:
                    print(f"⚠️  Could not remove {os.path.basename(lock_file)}: {e}")
                    print("   (This usually means the database is actually open)")
            
            if removed_count > 0:
                print(f"✅ Removed {removed_count} lock files safely")
            return True
        else:
            print("✅ Lock files left untouched (safe choice for production)")
            return False
    else:
        print("✅ No lock files found in source directory")
        return False

def clear_com_cache_safe():
    """Safely clear COM automation cache without affecting running processes."""
    print("🧹 Clearing COM cache (safe for production)...")
    
    try:
        import win32com.client
        import tempfile
        import shutil
        
        # Get current cache path
        cache_dir = win32com.client.gencache.GetGeneratePath()
        
        if os.path.exists(cache_dir):
            # Create backup first (safety measure)
            backup_dir = f"{cache_dir}_backup_{int(time.time())}"
            shutil.copytree(cache_dir, backup_dir, ignore_errors=True)
            
            # Clear cache
            shutil.rmtree(cache_dir, ignore_errors=True)
            print("✅ COM cache cleared (backup created)")
            print(f"   Backup location: {backup_dir}")
        else:
            print("ℹ️  COM cache directory not found")
    except Exception as e:
        print(f"⚠️  Could not clear COM cache: {e}")

def test_access_com_safe():
    """Test Access COM automation without interfering with running instances."""
    print("🧪 Testing Access COM automation (production-safe)...")
    
    try:
        import win32com.client
        
        # Try to create a NEW Access application instance
        # Use CreateObject instead of GetObject to avoid connecting to existing instances
        access_app = win32com.client.DispatchEx('Access.Application')  # DispatchEx creates new instance
        access_app.Visible = False  # Keep it invisible
        
        print("✅ New Access COM instance created successfully")
        
        # Test basic functionality without opening any databases
        version = getattr(access_app, 'Version', 'Unknown')
        print(f"✅ Access version: {version}")
        
        # Close cleanly
        access_app.Quit()
        access_app = None
        print("✅ New Access COM instance closed successfully")
        
        return True
    except Exception as e:
        print(f"❌ Access COM test failed: {e}")
        return False

def suggest_production_safe_solutions():
    """Suggest production-safe solutions for database lock issues."""
    print("\n💡 PRODUCTION-SAFE SOLUTIONS FOR DATABASE LOCK ISSUES:")
    print("=" * 60)
    
    print("\n1. 🔄 CONVERTER BUILT-IN RETRY MECHANISM:")
    print("   - The enhanced converter automatically retries locked databases")
    print("   - It waits between attempts and closes databases properly")
    print("   - This solves most lock issues without manual intervention")
    
    print("\n2. 🕒 SCHEDULE DURING OFF-PEAK HOURS:")
    print("   - Run conversions when databases are less likely to be in use")
    print("   - Early morning or late evening are typically safer")
    
    print("\n3. 📂 COPY-AND-CONVERT APPROACH:")
    print("   - Copy MDB files to a separate directory first")
    print("   - Run conversion on the copies, not the originals")
    print("   - This eliminates production interference")
    
    print("\n4. 🔍 CHECK CURRENT DATABASE USAGE:")
    print("   - Use network monitoring tools to see who's accessing databases")
    print("   - Coordinate with users before running conversions")
    
    print("\n5. 🛡️ INCREMENTAL APPROACH:")
    print("   - Convert databases one at a time")
    print("   - Use --no-progress-thread for minimal system impact")
    print("   - Monitor system resources during conversion")

def main():
    """Main function for production-safe database lock fixing."""
    print("🛡️  PRODUCTION-SAFE MS ACCESS LOCK CHECKER")
    print("=" * 50)
    print("This script is safe to run in production environments.")
    print("It will NOT kill processes or force-remove active lock files.")
    print("=" * 50)
    
    # Get source directory
    source_dir = None
    if len(sys.argv) > 1:
        source_dir = sys.argv[1]
    
    # Step 1: Check for running Access processes (non-destructive)
    has_running_access = check_access_processes()
    
    # Step 2: Check for lock files in source directory only
    has_locks = check_lock_files_in_source_only(source_dir)
    
    # Step 3: Clear COM cache safely
    clear_com_cache_safe()
    
    # Step 4: Wait for cleanup
    print("⏳ Waiting for cleanup to complete...")
    time.sleep(2)
    
    # Step 5: Test Access COM safely
    com_works = test_access_com_safe()
    
    # Step 6: Provide recommendations
    print("\n" + "=" * 60)
    if com_works and not has_running_access:
        print("✅ SYSTEM READY FOR CONVERSION")
        print("\nRecommendations:")
        print("- Run your conversion with the enhanced converter")
        print("- The built-in retry mechanism will handle any remaining locks")
        print("- Monitor the conversion logs for any issues")
        
        print("\nRun conversion with:")
        print("python access_com_converter.py \"path\\to\\mdb\\files\" --user username --password password")
        
    elif has_running_access:
        print("⚠️  ACCESS PROCESSES DETECTED")
        print("\nRecommendations:")
        print("- Coordinate with users currently using Access")
        print("- Consider running during off-peak hours")
        print("- Use copy-and-convert approach for safety")
        
    else:
        print("⚠️  POTENTIAL ISSUES DETECTED")
        print("\nRecommendations:")
        print("- Check Access installation")
        print("- Run as administrator if needed")
        print("- Consider restarting the system during maintenance window")
    
    # Always show production-safe solutions
    suggest_production_safe_solutions()
    
    print("\n🔗 For more help, see: DEPLOYMENT_CHECKLIST.md")
    input("\nPress Enter to exit...")

if __name__ == "__main__":
    main()
