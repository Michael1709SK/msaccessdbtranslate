#!/usr/bin/env python3
"""
MS Access to MySQL Database Converter - Runner Script

This script provides a convenient way to run the conversion process using
saved configuration or command-line arguments.
"""

import os
import sys
import json
import argparse
from pathlib import Path
from access_to_mysql_converter import AccessToMySQLConverter


def load_config(config_file="converter_config.json"):
    """Load configuration from JSON file."""
    config_path = Path(config_file)
    if config_path.exists():
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"Error loading config file: {e}")
            return None
    return None


def main():
    """Main function to run the converter with configuration or arguments."""
    parser = argparse.ArgumentParser(
        description="Run MS Access to MySQL conversion",
        epilog="If no arguments provided, will attempt to load from converter_config.json"
    )
    
    parser.add_argument("--source-dir", help="Directory containing MS Access database files")
    parser.add_argument("--host", help="MySQL host")
    parser.add_argument("--port", type=int, help="MySQL port")
    parser.add_argument("--user", help="MySQL username")
    parser.add_argument("--password", help="MySQL password")
    parser.add_argument("--log-dir", help="Directory for log files")
    parser.add_argument("--config", help="Configuration file path (default: converter_config.json)")
    
    args = parser.parse_args()
    
    # Determine configuration source
    config = None
    mysql_config = {}
    source_dir = None
    log_dir = "logs"
    
    if args.config or not any([args.source_dir, args.host, args.user, args.password]):
        # Try to load from configuration file
        config_file = args.config or "converter_config.json"
        config = load_config(config_file)
        
        if config:
            print(f"Using configuration from: {config_file}")
            source_dir = config.get('source_directory')
            mysql_config = config.get('mysql', {})
            log_dir = config.get('log_directory', 'logs')
        else:
            if not any([args.source_dir, args.host, args.user, args.password]):
                print("No configuration file found and no command-line arguments provided.")
                print("Please run: python config_setup.py setup")
                print("Or provide command-line arguments.")
                sys.exit(1)
    
    # Override with command-line arguments if provided
    if args.source_dir:
        source_dir = args.source_dir
    if args.host:
        mysql_config['host'] = args.host
    if args.port:
        mysql_config['port'] = args.port
    if args.user:
        mysql_config['user'] = args.user
    if args.password:
        mysql_config['password'] = args.password
    if args.log_dir:
        log_dir = args.log_dir
    
    # Validate required parameters
    if not source_dir:
        print("Error: Source directory is required")
        sys.exit(1)
    
    required_mysql_fields = ['host', 'user', 'password']
    missing_fields = [field for field in required_mysql_fields if not mysql_config.get(field)]
    
    if missing_fields:
        print(f"Error: Missing MySQL configuration: {', '.join(missing_fields)}")
        sys.exit(1)
    
    # Set defaults
    mysql_config.setdefault('host', 'localhost')
    mysql_config.setdefault('port', 3306)
    mysql_config['autocommit'] = False
    
    # Display configuration
    print("\n" + "=" * 60)
    print("MS Access to MySQL Conversion - Starting")
    print("=" * 60)
    print(f"Source Directory: {source_dir}")
    print(f"MySQL Host: {mysql_config['host']}:{mysql_config['port']}")
    print(f"MySQL User: {mysql_config['user']}")
    print(f"Log Directory: {log_dir}")
    print("=" * 60)
    
    # Confirm before proceeding
    if not config:  # Only ask for confirmation if not using config file
        response = input("\nProceed with conversion? (y/n): ").lower()
        if response != 'y':
            print("Conversion cancelled.")
            sys.exit(0)
    
    # Run the conversion
    try:
        converter = AccessToMySQLConverter(source_dir, mysql_config, log_dir)
        report = converter.run_conversion()
        
        # Display results
        print("\n" + "=" * 60)
        print("CONVERSION COMPLETED")
        print("=" * 60)
        
        stats = report['statistics']
        if stats['databases_failed'] == 0:
            print("✅ All databases converted successfully!")
        else:
            print(f"⚠️  Completed with {stats['databases_failed']} failures")
        
        print(f"\nSummary:")
        print(f"  Databases processed: {stats['databases_found']}")
        print(f"  Successfully converted: {stats['databases_converted']}")
        print(f"  Failed: {stats['databases_failed']}")
        print(f"  Tables converted: {stats['tables_converted']}")
        print(f"  Records migrated: {stats['records_migrated']}")
        
        success_rate = (stats['databases_converted'] / max(stats['databases_found'], 1)) * 100
        print(f"  Success rate: {success_rate:.1f}%")
        
        print(f"\nDetailed logs available in: {log_dir}")
        
        # Exit with appropriate code
        sys.exit(0 if stats['databases_failed'] == 0 else 1)
        
    except KeyboardInterrupt:
        print("\n\nConversion interrupted by user.")
        sys.exit(1)
    except Exception as e:
        print(f"\nUnexpected error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
