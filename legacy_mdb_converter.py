#!/usr/bin/env python3
"""
Alternative MS Access to MySQL converter for very old MDB files.
This version provides multiple connection methods and fallback options.
"""

import os
import sys
import logging
import subprocess
import tempfile
from pathlib import Path
from access_to_mysql_converter import AccessToMySQLConverter
import platform

try:
    import pyodbc
    import mysql.connector
except ImportError as e:
    print(f"Missing required package: {e}")
    sys.exit(1)


class LegacyAccessConverter(AccessToMySQLConverter):
    """Extended converter for very old Access databases (.mdb files)."""
    
    def __init__(self, source_dir: str, mysql_config: dict, log_dir: str = "logs"):
        super().__init__(source_dir, mysql_config, log_dir)
        self.conversion_methods = ['odbc', 'com', 'export']
    
    def check_access_installation(self) -> bool:
        """Check if Microsoft Access is installed on the system."""
        if platform.system() != "Windows":
            return False
        
        try:
            import winreg
            # Check for Access installation
            access_paths = [
                r"SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office",
                r"SOFTWARE\Microsoft\Office",
                r"SOFTWARE\WOW6432Node\Microsoft\Office"
            ]
            
            for base_path in access_paths:
                try:
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, base_path) as key:
                        i = 0
                        while True:
                            try:
                                subkey_name = winreg.EnumKey(key, i)
                                if subkey_name.replace(".", "").isdigit():
                                    try:
                                        access_path = f"{base_path}\\{subkey_name}\\Access"
                                        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, access_path):
                                            self.logger.info(f"Found Microsoft Access {subkey_name}")
                                            return True
                                    except FileNotFoundError:
                                        pass
                                i += 1
                            except OSError:
                                break
                except FileNotFoundError:
                    continue
            
            return False
        except ImportError:
            return False
    
    def convert_via_com(self, db_path: Path) -> bool:
        """Convert using COM automation (requires Access installation)."""
        try:
            self.logger.info(f"Attempting COM conversion for {db_path.name}")
            
            # This requires win32com which isn't in requirements
            try:
                import win32com.client
            except ImportError:
                self.logger.warning("win32com not available. Install with: pip install pywin32")
                return False
            
            # Create Access application
            access = win32com.client.Dispatch("Access.Application")
            access.Visible = False
            
            try:
                # Open the database
                access.OpenCurrentDatabase(str(db_path.absolute()))
                
                # Export each table to CSV, then import to MySQL
                db_name = self.sanitize_name(db_path.stem)
                temp_dir = Path(tempfile.mkdtemp())
                
                # Get table names
                table_names = []
                for table in access.CurrentData.AllTables:
                    if not table.Name.startswith("MSys"):
                        table_names.append(table.Name)
                
                self.logger.info(f"Found {len(table_names)} tables via COM")
                
                # Export each table
                for table_name in table_names:
                    try:
                        csv_file = temp_dir / f"{table_name}.csv"
                        access.DoCmd.TransferText(
                            TransferType=2,  # acExportDelim
                            TableName=table_name,
                            FileName=str(csv_file),
                            HasFieldNames=True
                        )
                        
                        # Import CSV to MySQL
                        self.import_csv_to_mysql(csv_file, db_name, table_name)
                        
                    except Exception as e:
                        self.logger.error(f"Failed to export table {table_name}: {e}")
                        continue
                
                return True
                
            finally:
                access.CloseCurrentDatabase()
                access.Quit()
                
        except Exception as e:
            self.logger.error(f"COM conversion failed: {e}")
            return False
    
    def import_csv_to_mysql(self, csv_file: Path, db_name: str, table_name: str) -> bool:
        """Import CSV file to MySQL."""
        try:
            import pandas as pd
            
            # Read CSV
            df = pd.read_csv(csv_file, encoding='utf-8', low_memory=False)
            
            if df.empty:
                return True
            
            # Connect to MySQL
            mysql_conn = self.connect_to_mysql()
            if not mysql_conn:
                return False
            
            cursor = mysql_conn.cursor()
            
            # Create database if not exists
            cursor.execute(f"CREATE DATABASE IF NOT EXISTS `{db_name}` CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci")
            mysql_conn.commit()
            
            # Infer table structure from CSV
            sanitized_table_name = self.sanitize_name(table_name)
            columns_sql = []
            
            for col in df.columns:
                sanitized_col = self.sanitize_name(col)
                # Simple type inference
                dtype = str(df[col].dtype)
                if 'int' in dtype:
                    mysql_type = 'INT'
                elif 'float' in dtype:
                    mysql_type = 'DOUBLE'
                elif 'datetime' in dtype:
                    mysql_type = 'DATETIME'
                else:
                    mysql_type = 'TEXT'
                
                columns_sql.append(f"`{sanitized_col}` {mysql_type}")
            
            # Create table
            create_sql = f"""
            CREATE TABLE IF NOT EXISTS `{db_name}`.`{sanitized_table_name}` (
                {',\n    '.join(columns_sql)}
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
            """
            
            cursor.execute(create_sql)
            mysql_conn.commit()
            
            # Insert data
            df.columns = [self.sanitize_name(col) for col in df.columns]
            df = df.where(pd.notnull(df), None)
            
            columns = ', '.join([f"`{col}`" for col in df.columns])
            placeholders = ', '.join(['%s'] * len(df.columns))
            insert_sql = f"INSERT INTO `{db_name}`.`{sanitized_table_name}` ({columns}) VALUES ({placeholders})"
            
            values = [tuple(row) for row in df.values]
            cursor.executemany(insert_sql, values)
            mysql_conn.commit()
            
            mysql_conn.close()
            
            self.logger.info(f"Successfully imported {len(df)} records from CSV")
            return True
            
        except Exception as e:
            self.logger.error(f"Failed to import CSV {csv_file}: {e}")
            return False
    
    def convert_via_mdb_tools(self, db_path: Path) -> bool:
        """Convert using mdb-tools (Linux/Mac/Windows with WSL)."""
        try:
            self.logger.info(f"Attempting mdb-tools conversion for {db_path.name}")
            
            # Check if mdb-tools is available
            result = subprocess.run(['mdb-tables', str(db_path)], 
                                  capture_output=True, text=True)
            
            if result.returncode != 0:
                self.logger.warning("mdb-tools not available")
                return False
            
            # Get table list
            tables = result.stdout.strip().split()
            self.logger.info(f"Found {len(tables)} tables via mdb-tools")
            
            db_name = self.sanitize_name(db_path.stem)
            temp_dir = Path(tempfile.mkdtemp())
            
            # Export each table
            for table_name in tables:
                if table_name.startswith('MSys'):
                    continue
                    
                try:
                    csv_file = temp_dir / f"{table_name}.csv"
                    
                    # Export table to CSV
                    with open(csv_file, 'w') as f:
                        result = subprocess.run(['mdb-export', str(db_path), table_name], 
                                              stdout=f, text=True)
                    
                    if result.returncode == 0:
                        self.import_csv_to_mysql(csv_file, db_name, table_name)
                    
                except Exception as e:
                    self.logger.error(f"Failed to export table {table_name}: {e}")
                    continue
            
            return True
            
        except Exception as e:
            self.logger.error(f"mdb-tools conversion failed: {e}")
            return False
    
    def convert_database(self, access_db_path: Path) -> bool:
        """Convert database using multiple methods as fallbacks."""
        self.logger.info(f"Starting legacy conversion of {access_db_path.name}")
        
        # Try ODBC first (original method)
        try:
            if super().convert_database(access_db_path):
                return True
        except Exception as e:
            self.logger.warning(f"ODBC method failed: {e}")
        
        # Try COM automation if Access is installed
        if platform.system() == "Windows" and self.check_access_installation():
            self.logger.info("Trying COM automation method...")
            if self.convert_via_com(access_db_path):
                return True
        
        # Try mdb-tools
        self.logger.info("Trying mdb-tools method...")
        if self.convert_via_mdb_tools(access_db_path):
            return True
        
        # All methods failed
        self.logger.error(f"All conversion methods failed for {access_db_path.name}")
        self.logger.error("Solutions:")
        self.logger.error("1. Install Microsoft Access Database Engine 2016")
        self.logger.error("2. Install Microsoft Access (for COM automation)")
        self.logger.error("3. Use mdb-tools (Linux/Mac/WSL)")
        self.logger.error("4. Convert manually using Access -> Export to CSV")
        
        return False


def main():
    """Main function using the legacy converter."""
    import argparse
    
    parser = argparse.ArgumentParser(description="Convert old MS Access (.mdb) databases to MySQL")
    parser.add_argument("source_dir", help="Directory containing MS Access database files")
    parser.add_argument("--host", default="localhost", help="MySQL host (default: localhost)")
    parser.add_argument("--port", type=int, default=3306, help="MySQL port (default: 3306)")
    parser.add_argument("--user", required=True, help="MySQL username")
    parser.add_argument("--password", required=True, help="MySQL password")
    parser.add_argument("--log-dir", default="logs", help="Directory for log files (default: logs)")
    
    args = parser.parse_args()
    
    mysql_config = {
        'host': args.host,
        'port': args.port,
        'user': args.user,
        'password': args.password,
        'autocommit': False
    }
    
    # Use legacy converter
    converter = LegacyAccessConverter(args.source_dir, mysql_config, args.log_dir)
    report = converter.run_conversion()
    
    if report['statistics']['databases_failed'] == 0:
        print("\n✅ All databases converted successfully!")
        sys.exit(0)
    else:
        print(f"\n⚠️  Conversion completed with {report['statistics']['databases_failed']} failures")
        sys.exit(1)


if __name__ == "__main__":
    main()
