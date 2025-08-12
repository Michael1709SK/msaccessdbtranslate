#!/usr/bin/env python3
"""
Sample usage script for the MS Access to MySQL converter.
This demonstrates how to use the converter programmatically.
"""

from pathlib import Path
from access_to_mysql_converter import AccessToMySQLConverter


def example_usage():
    """Example of how to use the converter programmatically."""
    
    # Configuration
    source_directory = r"C:\path\to\access\databases"  # Update this path
    log_directory = "logs"
    
    mysql_config = {
        'host': 'localhost',
        'port': 3306,
        'user': 'your_username',      # Update with your MySQL username
        'password': 'your_password',  # Update with your MySQL password
        'autocommit': False
    }
    
    print("MS Access to MySQL Converter - Example Usage")
    print("=" * 50)
    
    # Create converter instance
    converter = AccessToMySQLConverter(source_directory, mysql_config, log_directory)
    
    # Run the conversion
    try:
        report = converter.run_conversion()
        
        # Process the results
        stats = report['statistics']
        print(f"\nConversion Results:")
        print(f"  Databases found: {stats['databases_found']}")
        print(f"  Successfully converted: {stats['databases_converted']}")
        print(f"  Failed: {stats['databases_failed']}")
        print(f"  Tables converted: {stats['tables_converted']}")
        print(f"  Records migrated: {stats['records_migrated']}")
        
        if stats['databases_failed'] == 0:
            print("\n✅ All conversions completed successfully!")
        else:
            print(f"\n⚠️ {stats['databases_failed']} databases failed to convert")
            print("Check the log files for detailed error information.")
        
        return report
        
    except Exception as e:
        print(f"\nError during conversion: {e}")
        return None


if __name__ == "__main__":
    # Update the configuration above before running
    print("Please update the configuration in this script before running:")
    print("  - source_directory: Path to your Access databases")
    print("  - mysql_config: Your MySQL connection details")
    print("\nThen run: python example_usage.py")
    
    # Uncomment the line below after updating configuration
    # example_usage()
