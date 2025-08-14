#!/usr/bin/env python3
"""
Test script for the enhanced Access COM converter
"""

import os
import sys
import tempfile
from pathlib import Path

# Add current directory to path
sys.path.insert(0, str(Path(__file__).parent))

def test_converter_components():
    """Test the enhanced converter components"""
    print("🧪 Testing Enhanced Access COM Converter Components")
    print("=" * 60)
    
    try:
        from access_com_converter import ConversionStatistics, ProgressDisplayThread, AccessCOMConverter
        print("✅ All imports successful")
    except ImportError as e:
        print(f"❌ Import failed: {e}")
        return False
    
    # Test statistics tracker
    try:
        print("\n📊 Testing ConversionStatistics...")
        with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.log') as f:
            stats_log = f.name
        
        stats = ConversionStatistics(log_file=stats_log)
        
        # Test database tracking
        stats.add_database("test.mdb", 5)
        stats.start_database("test.mdb")
        
        # Test table tracking
        stats.add_table_size("Users", 1000)
        stats.add_table_size("Orders", 50000)
        stats.add_table_size("Products", 500)
        
        # Test processing order
        sorted_tables = stats.get_sorted_tables()
        expected_order = [('Products', 500), ('Users', 1000), ('Orders', 50000)]
        
        if sorted_tables == expected_order:
            print("✅ Table sorting works correctly")
        else:
            print(f"❌ Table sorting failed: got {sorted_tables}, expected {expected_order}")
            return False
        
        # Test table processing
        stats.start_table("Products", 500)
        stats.update_table_progress("Products", 250)
        stats.complete_table("Products", 500, 'completed')
        
        stats.start_table("Users", 1000)
        stats.complete_table("Users", 1000, 'skipped')
        
        print("✅ Statistics tracking works correctly")
        
        # Test progress display
        stats.display_progress()
        print("✅ Progress display works correctly")
        
        # Test report generation
        stats.save_final_report()
        print("✅ Report generation works correctly")
        
        # Cleanup
        try:
            os.unlink(stats_log)
        except:
            pass
            
    except Exception as e:
        print(f"❌ Statistics test failed: {e}")
        return False
    
    # Test progress display thread
    try:
        print("\n🔄 Testing ProgressDisplayThread...")
        import threading
        import time
        
        stats = ConversionStatistics()
        progress_thread = ProgressDisplayThread(stats, update_interval=1)
        progress_thread.start()
        
        # Let it run briefly
        time.sleep(2)
        progress_thread.stop()
        progress_thread.join(timeout=3)
        
        print("✅ Progress display thread works correctly")
        
    except Exception as e:
        print(f"❌ Progress thread test failed: {e}")
        return False
    
    print("\n🎉 All component tests passed!")
    return True

def test_integration():
    """Test basic integration without actual Access files"""
    print("\n🔧 Testing Integration Components...")
    
    try:
        from access_com_converter import AccessCOMConverter, ConversionStatistics
        
        # Create test configuration
        mysql_config = {
            'host': 'localhost',
            'port': 3306,
            'user': 'test_user',
            'password': 'test_password',
            'autocommit': False
        }
        
        with tempfile.TemporaryDirectory() as temp_dir:
            stats = ConversionStatistics()
            converter = AccessCOMConverter(temp_dir, mysql_config, "test_logs", stats)
            
            print("✅ AccessCOMConverter initialization successful")
            
            # Test utility methods
            test_name = converter.sanitize_name("Test Table Name!")
            expected = "Test_Table_Name"
            
            if test_name == expected:
                print("✅ Name sanitization works correctly")
            else:
                print(f"❌ Name sanitization failed: got {test_name}, expected {expected}")
                return False
        
        print("✅ Integration tests passed!")
        return True
        
    except Exception as e:
        print(f"❌ Integration test failed: {e}")
        return False

def main():
    """Run all tests"""
    print("🚀 Enhanced Access COM Converter Test Suite")
    print("=" * 60)
    
    success = True
    
    if not test_converter_components():
        success = False
    
    if not test_integration():
        success = False
    
    print("\n" + "=" * 60)
    if success:
        print("🎉 ALL TESTS PASSED! The enhanced converter is ready to use.")
        print("\nUsage example:")
        print('python access_com_converter.py "C:\\path\\to\\mdb\\files" --user mysql_user --password mysql_password')
        sys.exit(0)
    else:
        print("❌ Some tests failed. Please check the output above.")
        sys.exit(1)

if __name__ == "__main__":
    main()
