#!/usr/bin/env python3
"""
MS Access to MySQL converter using COM automation.
This works with installed Microsoft Access and is ideal for old MDB files.
"""

import os
import sys
import logging
import tempfile
import json
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Any, Optional
import re

try:
    import win32com.client
    import pandas as pd
    import mysql.connector
except ImportError as e:
    print(f"Missing required package: {e}")
    print("Please install: pip install pywin32 pandas mysql-connector-python")
    sys.exit(1)


class AccessCOMConverter:
    """Convert MS Access databases using COM automation (requires Access installation)."""
    
    def __init__(self, source_dir: str, mysql_config: Dict[str, str], log_dir: str = "logs"):
        self.source_dir = Path(source_dir)
        self.mysql_config = mysql_config
        self.log_dir = Path(log_dir)
        self.log_dir.mkdir(exist_ok=True)
        
        # Setup logging
        self.setup_logging()
        
        # Statistics tracking
        self.stats = {
            'databases_found': 0,
            'databases_converted': 0,
            'databases_failed': 0,
            'tables_converted': 0,
            'tables_failed': 0,
            'records_migrated': 0
        }
        
        # Access application object
        self.access_app = None
    
    def setup_logging(self):
        """Setup comprehensive logging system."""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_file = self.log_dir / f"access_com_converter_{timestamp}.log"
        
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_file, encoding='utf-8'),
                logging.StreamHandler(sys.stdout)
            ]
        )
        self.logger = logging.getLogger(__name__)
        self.logger.info("MS Access COM Converter initialized")
        self.logger.info(f"Source directory: {self.source_dir}")
        self.logger.info(f"Log file: {log_file}")
    
    def sanitize_name(self, name: str) -> str:
        """Sanitize database/table names for MySQL compatibility."""
        sanitized = re.sub(r'[^\w]', '_', name)
        if sanitized[0].isdigit():
            sanitized = f"db_{sanitized}"
        return sanitized.lower()[:64]
    
    def start_access(self) -> bool:
        """Start Microsoft Access application."""
        try:
            self.logger.info("Starting Microsoft Access...")
            self.access_app = win32com.client.Dispatch("Access.Application")
            self.access_app.Visible = False  # Keep Access hidden
            self.logger.info("✅ Microsoft Access started successfully")
            return True
        except Exception as e:
            self.logger.error(f"❌ Failed to start Microsoft Access: {e}")
            self.logger.error("Make sure Microsoft Access is properly installed")
            return False
    
    def close_access(self):
        """Close Microsoft Access application."""
        try:
            if self.access_app:
                try:
                    self.access_app.CloseCurrentDatabase()
                except:
                    pass
                self.access_app.Quit()
                self.access_app = None
                self.logger.info("Microsoft Access closed")
        except Exception as e:
            self.logger.warning(f"Error closing Access: {e}")
    
    def find_access_databases(self) -> List[Path]:
        """Find all MS Access database files in the source directory."""
        self.logger.info("Scanning for MS Access database files...")
        
        access_extensions = ['.mdb', '.accdb']
        databases = []
        
        for ext in access_extensions:
            pattern = f"**/*{ext}"
            found_files = list(self.source_dir.rglob(pattern))
            databases.extend(found_files)
            self.logger.info(f"Found {len(found_files)} {ext} files")
        
        self.stats['databases_found'] = len(databases)
        self.logger.info(f"Total databases found: {len(databases)}")
        
        for db in databases:
            self.logger.info(f"  - {db}")
        
        return databases
    
    def connect_to_mysql(self) -> Optional[mysql.connector.MySQLConnection]:
        """Connect to MySQL database."""
        try:
            conn = mysql.connector.connect(**self.mysql_config)
            self.logger.info("Connected to MySQL server")
            return conn
        except Exception as e:
            self.logger.error(f"Failed to connect to MySQL: {e}")
            return None
    
    def get_table_list(self, db_path: Path) -> List[str]:
        """Get list of user tables from Access database."""
        try:
            self.logger.info(f"Opening database: {db_path.name}")
            self.access_app.OpenCurrentDatabase(str(db_path.absolute()))
            
            # Get all table names
            tables = []
            for table in self.access_app.CurrentData.AllTables:
                table_name = table.Name
                # Skip system tables
                if not table_name.startswith("MSys") and not table_name.startswith("~"):
                    tables.append(table_name)
            
            self.logger.info(f"Found {len(tables)} user tables: {tables}")
            return tables
            
        except Exception as e:
            self.logger.error(f"Failed to get table list from {db_path}: {e}")
            return []
    
    def export_table_to_csv(self, table_name: str, temp_dir: Path) -> Optional[Path]:
        """Export Access table to CSV file."""
        try:
            csv_file = temp_dir / f"{table_name}.csv"
            
            self.logger.debug(f"Exporting {table_name} to CSV...")
            
            # Use Access DoCmd.TransferText to export as CSV
            self.access_app.DoCmd.TransferText(
                TransferType=2,  # acExportDelim (CSV export)
                TableName=table_name,
                FileName=str(csv_file.absolute()),
                HasFieldNames=True
            )
            
            if csv_file.exists() and csv_file.stat().st_size > 0:
                self.logger.debug(f"✅ Exported {table_name} to {csv_file.name}")
                return csv_file
            else:
                self.logger.warning(f"⚠️  Export of {table_name} resulted in empty file")
                return None
                
        except Exception as e:
            self.logger.error(f"❌ Failed to export {table_name}: {e}")
            return None
    
    def analyze_csv_structure(self, csv_file: Path) -> Dict[str, str]:
        """Analyze CSV file to determine MySQL column types."""
        try:
            # Read a sample of the CSV to infer types
            df_sample = pd.read_csv(csv_file, nrows=1000, encoding='utf-8')
            
            column_types = {}
            for col in df_sample.columns:
                col_clean = self.sanitize_name(col)
                
                # Analyze the column data to determine best MySQL type
                series = df_sample[col].dropna()
                
                if series.empty:
                    column_types[col_clean] = 'TEXT'
                    continue
                
                # Check for numeric types
                if pd.api.types.is_integer_dtype(series):
                    max_val = series.max()
                    if max_val < 128:
                        column_types[col_clean] = 'TINYINT'
                    elif max_val < 32768:
                        column_types[col_clean] = 'SMALLINT'
                    elif max_val < 2147483648:
                        column_types[col_clean] = 'INT'
                    else:
                        column_types[col_clean] = 'BIGINT'
                elif pd.api.types.is_float_dtype(series):
                    column_types[col_clean] = 'DOUBLE'
                elif pd.api.types.is_datetime64_any_dtype(series):
                    column_types[col_clean] = 'DATETIME'
                else:
                    # String data - determine appropriate size
                    max_length = series.astype(str).str.len().max()
                    if max_length <= 255:
                        column_types[col_clean] = f'VARCHAR({min(max_length + 50, 255)})'
                    else:
                        column_types[col_clean] = 'TEXT'
            
            return column_types
            
        except Exception as e:
            self.logger.error(f"Failed to analyze CSV structure: {e}")
            return {}
    
    def import_csv_to_mysql(self, csv_file: Path, db_name: str, table_name: str) -> int:
        """Import CSV file to MySQL database."""
        try:
            mysql_conn = self.connect_to_mysql()
            if not mysql_conn:
                return 0
            
            cursor = mysql_conn.cursor()
            
            # Create database if not exists
            cursor.execute(f"CREATE DATABASE IF NOT EXISTS `{db_name}` CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci")
            mysql_conn.commit()
            
            # Analyze CSV structure
            column_types = self.analyze_csv_structure(csv_file)
            
            if not column_types:
                self.logger.error(f"Could not determine structure for {csv_file}")
                mysql_conn.close()
                return 0
            
            # Create table
            columns_sql = [f"`{col}` {col_type}" for col, col_type in column_types.items()]
            create_sql = f"""
                CREATE TABLE IF NOT EXISTS `{db_name}`.`{table_name}` (
                    {',\n    '.join(columns_sql)}
                ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
            """
            
            cursor.execute(f"DROP TABLE IF EXISTS `{db_name}`.`{table_name}`")
            cursor.execute(create_sql)
            mysql_conn.commit()
            
            # Read and import CSV data
            df = pd.read_csv(csv_file, encoding='utf-8')
            
            if df.empty:
                mysql_conn.close()
                return 0
            
            # Clean column names to match our sanitized names
            df.columns = [self.sanitize_name(col) for col in df.columns]
            
            # Handle null values
            df = df.where(pd.notnull(df), None)
            
            # Insert data in batches
            columns = ', '.join([f"`{col}`" for col in df.columns])
            placeholders = ', '.join(['%s'] * len(df.columns))
            insert_sql = f"INSERT INTO `{db_name}`.`{table_name}` ({columns}) VALUES ({placeholders})"
            
            batch_size = 1000
            total_rows = len(df)
            
            for i in range(0, total_rows, batch_size):
                batch = df.iloc[i:i+batch_size]
                values = [tuple(row) for row in batch.values]
                cursor.executemany(insert_sql, values)
                mysql_conn.commit()
                
                self.logger.debug(f"Inserted batch {i//batch_size + 1} ({min(i+batch_size, total_rows)}/{total_rows} rows)")
            
            mysql_conn.close()
            self.logger.info(f"✅ Imported {total_rows} records to {db_name}.{table_name}")
            return total_rows
            
        except Exception as e:
            self.logger.error(f"❌ Failed to import {csv_file}: {e}")
            return 0
    
    def convert_database(self, db_path: Path) -> bool:
        """Convert a single Access database to MySQL."""
        db_name = self.sanitize_name(db_path.stem)
        self.logger.info(f"Starting conversion of {db_path.name} -> {db_name}")
        
        try:
            # Get table list
            tables = self.get_table_list(db_path)
            if not tables:
                self.logger.warning(f"No tables found in {db_path.name}")
                return True
            
            # Create temporary directory for CSV exports
            temp_dir = Path(tempfile.mkdtemp(prefix="access_export_"))
            self.logger.debug(f"Using temp directory: {temp_dir}")
            
            converted_tables = 0
            total_records = 0
            
            # Convert each table
            for table_name in tables:
                try:
                    sanitized_table_name = self.sanitize_name(table_name)
                    self.logger.info(f"Converting table: {table_name} -> {sanitized_table_name}")
                    
                    # Export to CSV
                    csv_file = self.export_table_to_csv(table_name, temp_dir)
                    if not csv_file:
                        self.stats['tables_failed'] += 1
                        continue
                    
                    # Import to MySQL
                    records = self.import_csv_to_mysql(csv_file, db_name, sanitized_table_name)
                    if records > 0:
                        converted_tables += 1
                        total_records += records
                        self.stats['tables_converted'] += 1
                    else:
                        self.stats['tables_failed'] += 1
                    
                    # Clean up CSV file
                    try:
                        csv_file.unlink()
                    except:
                        pass
                        
                except Exception as e:
                    self.logger.error(f"Failed to convert table {table_name}: {e}")
                    self.stats['tables_failed'] += 1
                    continue
            
            # Clean up temp directory
            try:
                temp_dir.rmdir()
            except:
                pass
            
            self.stats['records_migrated'] += total_records
            
            self.logger.info(f"Database conversion completed: {converted_tables}/{len(tables)} tables converted")
            self.logger.info(f"Total records migrated: {total_records}")
            
            return True
            
        except Exception as e:
            self.logger.error(f"Database conversion failed: {e}")
            return False
        finally:
            # Close the database
            try:
                self.access_app.CloseCurrentDatabase()
            except:
                pass
    
    def run_conversion(self) -> Dict[str, Any]:
        """Run the complete conversion process."""
        self.logger.info("Starting MS Access to MySQL conversion using COM automation")
        start_time = datetime.now()
        
        # Start Microsoft Access
        if not self.start_access():
            return self.get_summary_report(start_time)
        
        try:
            # Find all Access databases
            databases = self.find_access_databases()
            if not databases:
                self.logger.warning("No MS Access databases found")
                return self.get_summary_report(start_time)
            
            # Convert each database
            for db_path in databases:
                try:
                    self.logger.info(f"\n{'='*80}")
                    self.logger.info(f"Processing database: {db_path}")
                    self.logger.info(f"{'='*80}")
                    
                    if self.convert_database(db_path):
                        self.stats['databases_converted'] += 1
                        self.logger.info(f"✅ Successfully converted: {db_path.name}")
                    else:
                        self.stats['databases_failed'] += 1
                        self.logger.error(f"❌ Failed to convert: {db_path.name}")
                        
                except Exception as e:
                    self.stats['databases_failed'] += 1
                    self.logger.error(f"❌ Unexpected error processing {db_path}: {e}")
                    continue
            
            return self.get_summary_report(start_time)
            
        finally:
            # Always close Access
            self.close_access()
    
    def get_summary_report(self, start_time: datetime) -> Dict[str, Any]:
        """Generate and log summary report."""
        end_time = datetime.now()
        duration = end_time - start_time
        
        report = {
            'start_time': start_time.isoformat(),
            'end_time': end_time.isoformat(),
            'duration': str(duration),
            'statistics': self.stats.copy()
        }
        
        # Log summary
        self.logger.info(f"\n{'='*80}")
        self.logger.info("CONVERSION SUMMARY REPORT")
        self.logger.info(f"{'='*80}")
        self.logger.info(f"Start Time: {start_time}")
        self.logger.info(f"End Time: {end_time}")
        self.logger.info(f"Duration: {duration}")
        self.logger.info(f"\nStatistics:")
        self.logger.info(f"  Databases Found: {self.stats['databases_found']}")
        self.logger.info(f"  Databases Converted: {self.stats['databases_converted']}")
        self.logger.info(f"  Databases Failed: {self.stats['databases_failed']}")
        self.logger.info(f"  Tables Converted: {self.stats['tables_converted']}")
        self.logger.info(f"  Tables Failed: {self.stats['tables_failed']}")
        self.logger.info(f"  Records Migrated: {self.stats['records_migrated']}")
        
        success_rate = (self.stats['databases_converted'] / max(self.stats['databases_found'], 1)) * 100
        self.logger.info(f"  Success Rate: {success_rate:.1f}%")
        
        # Save report to JSON
        report_file = self.log_dir / f"com_conversion_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        with open(report_file, 'w', encoding='utf-8') as f:
            json.dump(report, f, indent=2, default=str)
        
        self.logger.info(f"\nDetailed report saved to: {report_file}")
        self.logger.info(f"{'='*80}")
        
        return report


def main():
    """Main function to run the COM converter."""
    import argparse
    
    parser = argparse.ArgumentParser(description="Convert MS Access databases to MySQL using COM automation")
    parser.add_argument("source_dir", help="Directory containing MS Access database files")
    parser.add_argument("--host", default="localhost", help="MySQL host (default: localhost)")
    parser.add_argument("--port", type=int, default=3306, help="MySQL port (default: 3306)")
    parser.add_argument("--user", required=True, help="MySQL username")
    parser.add_argument("--password", required=True, help="MySQL password")
    parser.add_argument("--log-dir", default="logs", help="Directory for log files (default: logs)")
    
    args = parser.parse_args()
    
    # MySQL configuration
    mysql_config = {
        'host': args.host,
        'port': args.port,
        'user': args.user,
        'password': args.password,
        'autocommit': False
    }
    
    # Create converter and run
    converter = AccessCOMConverter(args.source_dir, mysql_config, args.log_dir)
    report = converter.run_conversion()
    
    # Exit with appropriate code
    if report['statistics']['databases_failed'] == 0:
        print("\n✅ All databases converted successfully!")
        sys.exit(0)
    else:
        print(f"\n⚠️  Conversion completed with {report['statistics']['databases_failed']} failures")
        print("Check the log files for detailed error information")
        sys.exit(1)


if __name__ == "__main__":
    main()
