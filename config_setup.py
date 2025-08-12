#!/usr/bin/env python3
"""
MS Access to MySQL Database Converter - Configuration Script

This script provides an interactive setup for configuring the database conversion process.
It creates configuration files and validates connections before running the conversion.
"""

import os
import json
import getpass
import sys
from pathlib import Path
from typing import Dict, Any

try:
    import mysql.connector
    import pyodbc
except ImportError as e:
    print(f"Missing required package: {e}")
    print("Please install required packages:")
    print("pip install pyodbc mysql-connector-python pandas")
    sys.exit(1)


class ConverterConfig:
    """Configuration manager for the Access to MySQL converter."""
    
    def __init__(self):
        self.config_file = Path("converter_config.json")
        self.config = {}
    
    def interactive_setup(self):
        """Interactive setup for configuration."""
        print("=" * 60)
        print("MS Access to MySQL Converter - Configuration Setup")
        print("=" * 60)
        
        # Source directory setup
        self.setup_source_directory()
        
        # MySQL connection setup
        self.setup_mysql_connection()
        
        # Output directories setup
        self.setup_output_directories()
        
        # Advanced options
        self.setup_advanced_options()
        
        # Save configuration
        self.save_config()
        
        # Test connections
        if input("\nTest connections now? (y/n): ").lower() == 'y':
            self.test_connections()
        
        print("\n" + "=" * 60)
        print("Configuration completed successfully!")
        print("=" * 60)
        print(f"Configuration saved to: {self.config_file.absolute()}")
        print("\nTo run the conversion, use:")
        print("python run_converter.py")
        print("\nOr manually:")
        print("python access_to_mysql_converter.py <source_dir> --user <user> --password <password>")
    
    def setup_source_directory(self):
        """Setup source directory configuration."""
        print("\n1. Source Directory Configuration")
        print("-" * 40)
        
        while True:
            source_dir = input("Enter the directory containing MS Access databases: ").strip()
            if not source_dir:
                print("Please enter a valid directory path")
                continue
                
            source_path = Path(source_dir)
            if not source_path.exists():
                print(f"Directory does not exist: {source_dir}")
                if input("Create directory? (y/n): ").lower() == 'y':
                    source_path.mkdir(parents=True, exist_ok=True)
                    print(f"Created directory: {source_path.absolute()}")
                else:
                    continue
            
            self.config['source_directory'] = str(source_path.absolute())
            print(f"Source directory set to: {self.config['source_directory']}")
            break
    
    def setup_mysql_connection(self):
        """Setup MySQL connection configuration."""
        print("\n2. MySQL Connection Configuration")
        print("-" * 40)
        
        self.config['mysql'] = {}
        
        # Host
        host = input("MySQL Host (default: localhost): ").strip()
        self.config['mysql']['host'] = host if host else 'localhost'
        
        # Port
        while True:
            port = input("MySQL Port (default: 3306): ").strip()
            if not port:
                self.config['mysql']['port'] = 3306
                break
            try:
                self.config['mysql']['port'] = int(port)
                break
            except ValueError:
                print("Please enter a valid port number")
        
        # Username
        while True:
            username = input("MySQL Username: ").strip()
            if username:
                self.config['mysql']['user'] = username
                break
            print("Username is required")
        
        # Password
        while True:
            password = getpass.getpass("MySQL Password: ")
            if password:
                self.config['mysql']['password'] = password
                break
            print("Password is required")
        
        print(f"MySQL connection configured for {username}@{self.config['mysql']['host']}:{self.config['mysql']['port']}")
    
    def setup_output_directories(self):
        """Setup output directory configuration."""
        print("\n3. Output Directory Configuration")
        print("-" * 40)
        
        # Log directory
        log_dir = input("Log directory (default: logs): ").strip()
        self.config['log_directory'] = log_dir if log_dir else 'logs'
        
        # Backup directory for original databases
        backup_dir = input("Backup directory for original databases (optional): ").strip()
        if backup_dir:
            self.config['backup_directory'] = backup_dir
        
        print(f"Logs will be saved to: {self.config['log_directory']}")
        if backup_dir:
            print(f"Database backups will be saved to: {backup_dir}")
    
    def setup_advanced_options(self):
        """Setup advanced configuration options."""
        print("\n4. Advanced Options")
        print("-" * 40)
        
        # Batch size for data migration
        batch_size = input("Data migration batch size (default: 1000): ").strip()
        try:
            self.config['batch_size'] = int(batch_size) if batch_size else 1000
        except ValueError:
            self.config['batch_size'] = 1000
        
        # Include system tables
        include_system = input("Include system tables? (y/n, default: n): ").lower()
        self.config['include_system_tables'] = include_system == 'y'
        
        # Auto-create indexes
        create_indexes = input("Auto-create indexes based on Access indexes? (y/n, default: y): ").lower()
        self.config['create_indexes'] = create_indexes != 'n'
        
        # Encoding
        encoding = input("Character encoding (default: utf8mb4): ").strip()
        self.config['encoding'] = encoding if encoding else 'utf8mb4'
        
        print(f"Advanced options configured:")
        print(f"  Batch size: {self.config['batch_size']}")
        print(f"  Include system tables: {self.config['include_system_tables']}")
        print(f"  Create indexes: {self.config['create_indexes']}")
        print(f"  Encoding: {self.config['encoding']}")
    
    def save_config(self):
        """Save configuration to file."""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=2)
            print(f"\nConfiguration saved to: {self.config_file.absolute()}")
        except Exception as e:
            print(f"Error saving configuration: {e}")
    
    def load_config(self) -> Dict[str, Any]:
        """Load configuration from file."""
        try:
            if self.config_file.exists():
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    self.config = json.load(f)
                print(f"Configuration loaded from: {self.config_file.absolute()}")
                return self.config
            else:
                print("No configuration file found. Please run setup first.")
                return {}
        except Exception as e:
            print(f"Error loading configuration: {e}")
            return {}
    
    def test_connections(self):
        """Test MySQL connection."""
        print("\n" + "=" * 40)
        print("Testing Connections")
        print("=" * 40)
        
        # Test MySQL connection
        try:
            mysql_config = self.config.get('mysql', {})
            conn = mysql.connector.connect(**mysql_config)
            cursor = conn.cursor()
            cursor.execute("SELECT VERSION()")
            version = cursor.fetchone()[0]
            conn.close()
            print(f"✅ MySQL connection successful - Version: {version}")
        except Exception as e:
            print(f"❌ MySQL connection failed: {e}")
        
        # Test Access driver availability
        try:
            drivers = [x for x in pyodbc.drivers() if 'Access' in x or 'Microsoft Access Driver' in x]
            if drivers:
                print(f"✅ MS Access drivers available: {drivers}")
            else:
                print("⚠️  No MS Access drivers found. You may need to install Microsoft Access Database Engine.")
        except Exception as e:
            print(f"❌ Error checking Access drivers: {e}")
    
    def display_current_config(self):
        """Display current configuration."""
        if not self.config:
            print("No configuration loaded.")
            return
        
        print("\nCurrent Configuration:")
        print("-" * 40)
        print(f"Source Directory: {self.config.get('source_directory', 'Not set')}")
        mysql_cfg = self.config.get('mysql', {})
        print(f"MySQL Host: {mysql_cfg.get('host', 'Not set')}")
        print(f"MySQL Port: {mysql_cfg.get('port', 'Not set')}")
        print(f"MySQL User: {mysql_cfg.get('user', 'Not set')}")
        print(f"Log Directory: {self.config.get('log_directory', 'Not set')}")
        print(f"Batch Size: {self.config.get('batch_size', 'Not set')}")


def main():
    """Main function for configuration setup."""
    config_manager = ConverterConfig()
    
    if len(sys.argv) > 1:
        command = sys.argv[1].lower()
        
        if command == 'setup':
            config_manager.interactive_setup()
        elif command == 'test':
            config_manager.load_config()
            config_manager.test_connections()
        elif command == 'show':
            config_manager.load_config()
            config_manager.display_current_config()
        else:
            print("Unknown command. Use: setup, test, or show")
    else:
        print("MS Access to MySQL Converter - Configuration Tool")
        print("\nUsage:")
        print("  python config_setup.py setup  - Run interactive setup")
        print("  python config_setup.py test   - Test connections")
        print("  python config_setup.py show   - Show current configuration")
        print("\nFor first-time setup, run:")
        print("  python config_setup.py setup")


if __name__ == "__main__":
    main()
