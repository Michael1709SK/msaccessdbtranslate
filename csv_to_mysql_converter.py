#!/usr/bin/env python3
"""
Manual CSV to MySQL converter.
This script helps when all ODBC methods fail for old MDB files.

Steps:
1. Open your .mdb files in Microsoft Access
2. Export each table as CSV (File → Export → Text File)
3. Run this script to import all CSV files to MySQL
"""

import os
import sys
import logging
from pathlib import Path
from datetime import datetime
import re

try:
    import pandas as pd
    import mysql.connector
except ImportError as e:
    print(f"Missing required package: {e}")
    print("pip install pandas mysql-connector-python")
    sys.exit(1)


class CSVToMySQLConverter:
    """Converts CSV files exported from Access to MySQL."""
    
    def __init__(self, csv_dir: str, mysql_config: dict, database_name: str = None):
        self.csv_dir = Path(csv_dir)
        self.mysql_config = mysql_config
        self.database_name = database_name or self.sanitize_name(self.csv_dir.name)
        
        # Setup logging
        self.setup_logging()
    
    def setup_logging(self):
        """Setup logging."""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_file = f"csv_to_mysql_{timestamp}.log"
        
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_file),
                logging.StreamHandler(sys.stdout)
            ]
        )
        self.logger = logging.getLogger(__name__)
    
    def sanitize_name(self, name: str) -> str:
        """Sanitize names for MySQL."""
        sanitized = re.sub(r'[^\w]', '_', name)
        if sanitized[0].isdigit():
            sanitized = f"db_{sanitized}"
        return sanitized.lower()[:64]
    
    def detect_delimiter(self, file_path: Path) -> str:
        """Detect CSV delimiter."""
        with open(file_path, 'r', encoding='utf-8') as f:
            first_line = f.readline()
            
        if first_line.count('\t') > first_line.count(','):
            return '\t'
        return ','
    
    def infer_mysql_type(self, series) -> str:
        """Infer MySQL type from pandas series."""
        dtype = str(series.dtype)
        
        # Check for integers
        if 'int' in dtype:
            max_val = series.max() if not series.empty else 0
            if max_val < 128:
                return 'TINYINT'
            elif max_val < 32768:
                return 'SMALLINT'
            elif max_val < 2147483648:
                return 'INT'
            else:
                return 'BIGINT'
        
        # Check for floats
        if 'float' in dtype:
            return 'DOUBLE'
        
        # Check for dates
        if 'datetime' in dtype:
            return 'DATETIME'
        
        # Check string length
        if dtype == 'object':
            max_len = series.astype(str).str.len().max() if not series.empty else 50
            if max_len <= 255:
                return f'VARCHAR({min(max_len + 50, 255)})'
            else:
                return 'TEXT'
        
        return 'TEXT'
    
    def convert_csv_file(self, csv_file: Path) -> bool:
        """Convert a single CSV file to MySQL table."""
        try:
            table_name = self.sanitize_name(csv_file.stem)
            self.logger.info(f"Converting {csv_file.name} -> {table_name}")
            
            # Detect delimiter
            delimiter = self.detect_delimiter(csv_file)
            
            # Read CSV with multiple encoding attempts
            encodings = ['utf-8', 'latin-1', 'cp1252', 'utf-16']
            df = None
            
            for encoding in encodings:
                try:
                    df = pd.read_csv(csv_file, delimiter=delimiter, encoding=encoding, low_memory=False)
                    self.logger.info(f"Successfully read with {encoding} encoding")
                    break
                except UnicodeDecodeError:
                    continue
            
            if df is None:
                self.logger.error(f"Could not read {csv_file} with any encoding")
                return False
            
            if df.empty:
                self.logger.warning(f"{csv_file.name} is empty")
                return True
            
            # Clean column names
            df.columns = [self.sanitize_name(col) for col in df.columns]
            
            # Connect to MySQL
            conn = mysql.connector.connect(**self.mysql_config)
            cursor = conn.cursor()
            
            # Create database
            cursor.execute(f"CREATE DATABASE IF NOT EXISTS `{self.database_name}`")
            conn.commit()
            
            # Create table
            columns_sql = []
            for col in df.columns:
                mysql_type = self.infer_mysql_type(df[col])
                columns_sql.append(f"`{col}` {mysql_type}")
            
            create_sql = f"""
            CREATE TABLE `{self.database_name}`.`{table_name}` (
                {',\n    '.join(columns_sql)}
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
            """
            
            cursor.execute(f"DROP TABLE IF EXISTS `{self.database_name}`.`{table_name}`")
            cursor.execute(create_sql)
            conn.commit()
            
            # Insert data
            df = df.where(pd.notnull(df), None)
            
            columns = ', '.join([f"`{col}`" for col in df.columns])
            placeholders = ', '.join(['%s'] * len(df.columns))
            insert_sql = f"INSERT INTO `{self.database_name}`.`{table_name}` ({columns}) VALUES ({placeholders})"
            
            # Insert in batches
            batch_size = 1000
            total_rows = len(df)
            
            for i in range(0, total_rows, batch_size):
                batch = df.iloc[i:i+batch_size]
                values = [tuple(row) for row in batch.values]
                cursor.executemany(insert_sql, values)
                conn.commit()
                
                self.logger.info(f"Inserted batch {i//batch_size + 1} ({min(i+batch_size, total_rows)}/{total_rows} rows)")
            
            conn.close()
            self.logger.info(f"✅ Successfully converted {csv_file.name} ({total_rows} rows)")
            return True
            
        except Exception as e:
            self.logger.error(f"❌ Failed to convert {csv_file.name}: {e}")
            return False
    
    def convert_all_csv_files(self):
        """Convert all CSV files in the directory."""
        csv_files = list(self.csv_dir.glob('*.csv'))
        
        if not csv_files:
            self.logger.error(f"No CSV files found in {self.csv_dir}")
            return
        
        self.logger.info(f"Found {len(csv_files)} CSV files to convert")
        
        successful = 0
        for csv_file in csv_files:
            if self.convert_csv_file(csv_file):
                successful += 1
        
        self.logger.info(f"Conversion completed: {successful}/{len(csv_files)} files successful")


def main():
    """Main function."""
    import argparse
    
    parser = argparse.ArgumentParser(
        description="Convert CSV files (exported from Access) to MySQL",
        epilog="First export your Access tables to CSV, then run this script"
    )
    
    parser.add_argument("csv_dir", help="Directory containing CSV files")
    parser.add_argument("--database-name", help="MySQL database name (default: directory name)")
    parser.add_argument("--host", default="localhost", help="MySQL host")
    parser.add_argument("--port", type=int, default=3306, help="MySQL port")
    parser.add_argument("--user", required=True, help="MySQL username")
    parser.add_argument("--password", required=True, help="MySQL password")
    
    args = parser.parse_args()
    
    mysql_config = {
        'host': args.host,
        'port': args.port,
        'user': args.user,
        'password': args.password,
        'autocommit': False
    }
    
    converter = CSVToMySQLConverter(args.csv_dir, mysql_config, args.database_name)
    converter.convert_all_csv_files()


if __name__ == "__main__":
    print("CSV to MySQL Converter")
    print("=" * 40)
    print("INSTRUCTIONS:")
    print("1. Open your .mdb files in Microsoft Access")
    print("2. For each table: File → Export → Text File")
    print("3. Save as CSV with headers")
    print("4. Put all CSV files in one directory")
    print("5. Run this script pointing to that directory")
    print("=" * 40)
    print()
    
    main()
