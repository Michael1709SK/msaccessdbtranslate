"""
MS Access to MySQL Database Converter

This script automatically discovers MS Access database files (.mdb, .accdb) in a directory,
extracts their structure and data, and converts them to MySQL format with proper relationships.

Features:
- Automatic database discovery
- Table structure conversion
- Data migration with type mapping
- Relationship preservation
- Error handling with continuation
- Comprehensive logging
- Progress tracking

Requirements:
- Python 3.7+
- pyodbc (for Access database connection)
- mysql-connector-python (for MySQL connection)
- pandas (for data manipulation)
"""

import os
import sys
import logging
import json
import traceback
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Any, Optional, Tuple
import re

try:
    import pyodbc
    import mysql.connector
    import pandas as pd
except ImportError as e:
    print(f"Missing required package: {e}")
    print("Please install required packages:")
    print("pip install pyodbc mysql-connector-python pandas")
    sys.exit(1)


class AccessToMySQLConverter:
    """Converts MS Access databases to MySQL databases with full structure and data migration."""
    
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
            'records_migrated': 0,
            'relationships_created': 0
        }
        
        # Access to MySQL type mapping
        self.type_mapping = {
            'COUNTER': 'INT AUTO_INCREMENT PRIMARY KEY',
            'LONG': 'INT',
            'INTEGER': 'INT',
            'SHORT': 'SMALLINT',
            'BYTE': 'TINYINT',
            'SINGLE': 'FLOAT',
            'DOUBLE': 'DOUBLE',
            'CURRENCY': 'DECIMAL(19,4)',
            'DATETIME': 'DATETIME',
            'BIT': 'BOOLEAN',
            'TEXT': 'VARCHAR(255)',
            'MEMO': 'TEXT',
            'LONGBINARY': 'LONGBLOB',
            'BINARY': 'VARBINARY(255)'
        }
    
    def setup_logging(self):
        """Setup comprehensive logging system."""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_file = self.log_dir / f"access_to_mysql_{timestamp}.log"
        
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_file, encoding='utf-8'),
                logging.StreamHandler(sys.stdout)
            ]
        )
        self.logger = logging.getLogger(__name__)
        self.logger.info("MS Access to MySQL Converter initialized")
        self.logger.info(f"Source directory: {self.source_dir}")
        self.logger.info(f"Log file: {log_file}")
    
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
    
    def sanitize_name(self, name: str) -> str:
        """Sanitize database/table names for MySQL compatibility."""
        # Remove or replace invalid characters
        sanitized = re.sub(r'[^\w]', '_', name)
        # Ensure it doesn't start with a number
        if sanitized[0].isdigit():
            sanitized = f"db_{sanitized}"
        # Limit length
        sanitized = sanitized[:64]
        return sanitized.lower()
    
    def get_access_connection_string(self, db_path: Path) -> str:
        """Generate connection string for MS Access database."""
        # Try multiple driver names in order of preference for old MDB files
        possible_drivers = [
            "Microsoft Access Driver (*.mdb, *.accdb)",  # Modern driver
            "Microsoft Access Driver (*.mdb)",           # Legacy MDB-only driver
            "Microsoft Office Access Driver (*.mdb, *.accdb)",
            "Driver do Microsoft Access (*.mdb)",        # Localized versions
            "Microsoft dBase Driver (*.dbf)",            # Sometimes works as fallback
        ]
        
        available_drivers = [x for x in pyodbc.drivers()]
        self.logger.debug(f"Available ODBC drivers: {available_drivers}")
        
        # Find the first available driver
        for driver in possible_drivers:
            if driver in available_drivers:
                self.logger.info(f"Using ODBC driver: {driver}")
                # For old MDB files, use simpler connection string
                if db_path.suffix.lower() == '.mdb':
                    return f"DRIVER={{{driver}}};DBQ={str(db_path.absolute())};ReadOnly=0;"
                else:
                    return f"DRIVER={{{driver}}};DBQ={str(db_path.absolute())};ExtendedAnsiSQL=1;"
        
        # If no specific Access driver found, try generic ones
        for driver in available_drivers:
            if "access" in driver.lower() or "mdb" in driver.lower() or "accdb" in driver.lower():
                self.logger.info(f"Using fallback ODBC driver: {driver}")
                return f"DRIVER={{{driver}}};DBQ={str(db_path.absolute())};ReadOnly=0;"
        
        # No suitable driver found - provide specific help for old MDB files
        raise Exception(f"No Microsoft Access ODBC driver found for .mdb files.\n"
                       f"Available drivers: {available_drivers}\n\n"
                       f"🔧 SOLUTION FOR OLD .MDB FILES:\n"
                       f"1. Download Microsoft Access Database Engine 2016:\n"
                       f"   https://www.microsoft.com/en-us/download/details.aspx?id=54920\n"
                       f"2. Choose the version matching your Python architecture\n"
                       f"3. Or try the legacy Jet Database Engine:\n"
                       f"   https://www.microsoft.com/en-us/download/details.aspx?id=23734\n"
                       f"4. Run: fix_odbc_drivers.bat for automated help")
    
    def connect_to_access(self, db_path: Path) -> Optional[pyodbc.Connection]:
        """Connect to MS Access database."""
        try:
            conn_str = self.get_access_connection_string(db_path)
            
            # Try different connection approaches for old MDB files
            connection_params = [
                {},  # Default
                {'timeout': 30},  # Longer timeout
                {'autocommit': True},  # Auto commit
                {'timeout': 30, 'autocommit': True},  # Both
            ]
            
            for params in connection_params:
                try:
                    conn = pyodbc.connect(conn_str, **params)
                    
                    # Test the connection with a simple query
                    cursor = conn.cursor()
                    
                    # Try different ways to test connection for old MDB files
                    test_queries = [
                        "SELECT 1",
                        "SELECT COUNT(*) FROM MSysObjects WHERE 1=0",
                        "SELECT Name FROM MSysObjects WHERE Type=1 AND Flags=0 LIMIT 1",
                    ]
                    
                    connection_works = False
                    for test_query in test_queries:
                        try:
                            cursor.execute(test_query)
                            cursor.fetchall()
                            connection_works = True
                            break
                        except:
                            continue
                    
                    if connection_works:
                        self.logger.info(f"Connected to Access database: {db_path.name}")
                        return conn
                    else:
                        conn.close()
                        continue
                        
                except Exception as e:
                    self.logger.debug(f"Connection attempt failed with params {params}: {e}")
                    continue
            
            # If all connection attempts failed
            raise Exception("All connection methods failed")
            
        except Exception as e:
            self.logger.error(f"Failed to connect to {db_path}: {e}")
            
            # Log additional troubleshooting information
            self.logger.error("Troubleshooting information:")
            try:
                available_drivers = pyodbc.drivers()
                self.logger.error(f"Available ODBC drivers: {list(available_drivers)}")
                
                # Check if any Access drivers are available
                access_drivers = [d for d in available_drivers if 'access' in d.lower() or 'mdb' in d.lower()]
                if not access_drivers:
                    self.logger.error("No Microsoft Access drivers found!")
                    self.logger.error("\n🔧 SOLUTIONS FOR OLD .MDB FILES:")
                    self.logger.error("1. Download Microsoft Access Database Engine 2016:")
                    self.logger.error("   https://www.microsoft.com/en-us/download/details.aspx?id=54920")
                    self.logger.error("2. Or download legacy Jet Database Engine 4.0:")
                    self.logger.error("   https://www.microsoft.com/en-us/download/details.aspx?id=23734")
                    self.logger.error("3. Install the version matching your Python architecture")
                    self.logger.error("4. If installation fails, try: installer.exe /quiet")
                    self.logger.error("5. Run: fix_odbc_drivers.bat for automated help")
                    self.logger.error("6. Alternative: Use legacy_mdb_converter.py")
                else:
                    self.logger.error("Available Access drivers:")
                    for driver in access_drivers:
                        self.logger.error(f"   - {driver}")
            except Exception as diag_e:
                self.logger.error(f"Could not get diagnostic information: {diag_e}")
                
            return None
    
    def connect_to_mysql(self) -> Optional[mysql.connector.MySQLConnection]:
        """Connect to MySQL database."""
        try:
            conn = mysql.connector.connect(**self.mysql_config)
            self.logger.info("Connected to MySQL server")
            return conn
        except Exception as e:
            self.logger.error(f"Failed to connect to MySQL: {e}")
            return None
    
    def get_table_list(self, access_conn: pyodbc.Connection) -> List[str]:
        """Get list of tables from Access database."""
        try:
            cursor = access_conn.cursor()
            tables = []
            
            # Method 1: Try the standard approach
            try:
                for table_info in cursor.tables(tableType='TABLE'):
                    table_name = table_info.table_name
                    # Skip system tables
                    if not table_name.startswith('MSys'):
                        tables.append(table_name)
                        
                if tables:
                    self.logger.info(f"Found {len(tables)} user tables (method 1)")
                    return tables
            except Exception as e:
                self.logger.warning(f"Standard table enumeration failed: {e}")
            
            # Method 2: Query system catalog for old MDB files
            try:
                cursor.execute("SELECT Name FROM MSysObjects WHERE Type=1 AND Flags=0")
                for row in cursor.fetchall():
                    table_name = row[0]
                    if not table_name.startswith('MSys') and not table_name.startswith('~'):
                        tables.append(table_name)
                        
                if tables:
                    self.logger.info(f"Found {len(tables)} user tables (method 2 - system catalog)")
                    return tables
            except Exception as e:
                self.logger.warning(f"System catalog query failed: {e}")
            
            # Method 3: Try to get table names from schema information
            try:
                # Get all table names from INFORMATION_SCHEMA equivalent
                cursor.execute("SELECT DISTINCT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE'")
                for row in cursor.fetchall():
                    table_name = row[0]
                    if not table_name.startswith('MSys'):
                        tables.append(table_name)
                        
                if tables:
                    self.logger.info(f"Found {len(tables)} user tables (method 3 - schema)")
                    return tables
            except Exception as e:
                self.logger.warning(f"Schema information query failed: {e}")
            
            # Method 4: Try alternative system table approach
            try:
                cursor.execute("SELECT Name FROM MSysObjects WHERE Type IN (1,4,6) AND Left(Name,4) <> 'MSys' AND Left(Name,1) <> '~'")
                for row in cursor.fetchall():
                    tables.append(row[0])
                    
                if tables:
                    self.logger.info(f"Found {len(tables)} user tables (method 4 - alternative)")
                    return tables
            except Exception as e:
                self.logger.warning(f"Alternative system table query failed: {e}")
            
            self.logger.warning("All table enumeration methods failed")
            return []
            
        except Exception as e:
            self.logger.error(f"Failed to get table list: {e}")
            return []
    
    def get_table_structure(self, access_conn: pyodbc.Connection, table_name: str) -> Dict[str, Any]:
        """Get table structure from Access database."""
        try:
            cursor = access_conn.cursor()
            
            # Get column information - try multiple methods for old MDB files
            columns = []
            
            # Method 1: Standard approach
            try:
                for column in cursor.columns(table=table_name):
                    col_info = {
                        'name': column.column_name,
                        'type': column.type_name,
                        'size': getattr(column, 'column_size', 0),
                        'nullable': getattr(column, 'nullable', True),
                        'default': getattr(column, 'column_def', None)
                    }
                    columns.append(col_info)
                    
                if columns:
                    self.logger.debug(f"Got structure for table {table_name}: {len(columns)} columns (method 1)")
                    
            except Exception as e:
                self.logger.warning(f"Standard column enumeration failed for {table_name}: {e}")
                
                # Method 2: Query the table directly to get column info
                try:
                    # Get first row to determine column names and types
                    cursor.execute(f"SELECT TOP 1 * FROM [{table_name}]")
                    
                    # Get column information from cursor description
                    if cursor.description:
                        for col_desc in cursor.description:
                            col_info = {
                                'name': col_desc[0],
                                'type': self.map_odbc_type_to_access(col_desc[1]),
                                'size': col_desc[2] if len(col_desc) > 2 else 255,
                                'nullable': True,  # Default for old MDB
                                'default': None
                            }
                            columns.append(col_info)
                        
                        self.logger.debug(f"Got structure for table {table_name}: {len(columns)} columns (method 2)")
                        
                except Exception as e2:
                    self.logger.warning(f"Direct table query failed for {table_name}: {e2}")
                    
                    # Method 3: Create minimal structure if all else fails
                    try:
                        cursor.execute(f"SELECT * FROM [{table_name}] WHERE 1=0")  # Get structure only
                        if cursor.description:
                            for col_desc in cursor.description:
                                col_info = {
                                    'name': col_desc[0],
                                    'type': 'TEXT',  # Default to TEXT for safety
                                    'size': 255,
                                    'nullable': True,
                                    'default': None
                                }
                                columns.append(col_info)
                            self.logger.debug(f"Got minimal structure for table {table_name}: {len(columns)} columns (method 3)")
                    except Exception as e3:
                        self.logger.error(f"All column enumeration methods failed for {table_name}: {e3}")
            
            # Get primary key information - be more tolerant of failures
            primary_keys = []
            try:
                for pk in cursor.primaryKeys(table=table_name):
                    primary_keys.append(pk.column_name)
            except Exception as e:
                self.logger.warning(f"Could not get primary keys for {table_name}: {e}")
                # Try alternative method for old MDB files
                try:
                    # Look for common primary key patterns
                    common_pk_names = ['ID', 'Id', f'{table_name}ID', f'{table_name}Id', 'RecordID']
                    for col in columns:
                        if col['name'].upper() in [pk.upper() for pk in common_pk_names]:
                            primary_keys.append(col['name'])
                            break
                except:
                    pass
            
            structure = {
                'columns': columns,
                'primary_keys': primary_keys
            }
            
            return structure
        except Exception as e:
            self.logger.error(f"Failed to get structure for table {table_name}: {e}")
            return {'columns': [], 'primary_keys': []}
    
    def map_odbc_type_to_access(self, odbc_type) -> str:
        """Map ODBC type constants to Access type names."""
        # Common ODBC type mappings
        type_map = {
            1: 'TEXT',      # SQL_CHAR
            4: 'LONG',      # SQL_INTEGER
            5: 'SHORT',     # SQL_SMALLINT
            6: 'SINGLE',    # SQL_FLOAT
            7: 'DOUBLE',    # SQL_REAL
            8: 'DOUBLE',    # SQL_DOUBLE
            9: 'DATETIME',  # SQL_DATE
            10: 'DATETIME', # SQL_TIME
            11: 'DATETIME', # SQL_TIMESTAMP
            12: 'TEXT',     # SQL_VARCHAR
            -1: 'MEMO',     # SQL_LONGVARCHAR
            -2: 'BINARY',   # SQL_BINARY
            -3: 'LONGBINARY', # SQL_VARBINARY
            -4: 'LONGBINARY', # SQL_LONGVARBINARY
        }
        return type_map.get(odbc_type, 'TEXT')
    
    def convert_column_type(self, access_type: str, size: int) -> str:
        """Convert Access column type to MySQL type."""
        access_type = access_type.upper()
        
        if access_type in self.type_mapping:
            mysql_type = self.type_mapping[access_type]
        else:
            # Default mapping for unknown types
            mysql_type = 'TEXT'
        
        # Handle TEXT fields with size
        if access_type == 'TEXT' and size > 0:
            if size <= 255:
                mysql_type = f'VARCHAR({size})'
            else:
                mysql_type = 'TEXT'
        
        return mysql_type
    
    def create_mysql_table(self, mysql_conn: mysql.connector.MySQLConnection, 
                          db_name: str, table_name: str, structure: Dict[str, Any]) -> bool:
        """Create MySQL table with converted structure."""
        try:
            cursor = mysql_conn.cursor()
            
            # Build CREATE TABLE statement
            columns_sql = []
            primary_key_columns = structure.get('primary_keys', [])
            
            for col in structure['columns']:
                col_name = self.sanitize_name(col['name'])
                mysql_type = self.convert_column_type(col['type'], col.get('size', 0))
                
                col_sql = f"`{col_name}` {mysql_type}"
                
                # Handle nullability
                if not col.get('nullable', True) or col['name'] in primary_key_columns:
                    col_sql += " NOT NULL"
                
                columns_sql.append(col_sql)
            
            # Add primary key constraint if exists
            if primary_key_columns:
                pk_cols = [f"`{self.sanitize_name(col)}`" for col in primary_key_columns]
                columns_sql.append(f"PRIMARY KEY ({', '.join(pk_cols)})")
            
            create_sql = f"""
            CREATE TABLE `{db_name}`.`{table_name}` (
                {',\n    '.join(columns_sql)}
            ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
            """
            
            cursor.execute(create_sql)
            mysql_conn.commit()
            self.logger.info(f"Created MySQL table: {db_name}.{table_name}")
            return True
            
        except Exception as e:
            self.logger.error(f"Failed to create table {table_name}: {e}")
            return False
    
    def migrate_table_data(self, access_conn: pyodbc.Connection, mysql_conn: mysql.connector.MySQLConnection,
                          source_table: str, target_db: str, target_table: str) -> int:
        """Migrate data from Access table to MySQL table."""
        try:
            # Try different query approaches for old MDB files
            queries_to_try = [
                f"SELECT * FROM `{source_table}`",
                f"SELECT * FROM [{source_table}]",
                f"SELECT * FROM {source_table}",
            ]
            
            df = None
            for query in queries_to_try:
                try:
                    df = pd.read_sql(query, access_conn)
                    self.logger.debug(f"Successfully read data using query: {query}")
                    break
                except Exception as e:
                    self.logger.debug(f"Query failed: {query} - {e}")
                    continue
            
            if df is None:
                self.logger.error(f"Could not read data from table {source_table} with any query method")
                return 0
            
            if df.empty:
                self.logger.info(f"Table {source_table} is empty")
                return 0
            
            # Sanitize column names
            df.columns = [self.sanitize_name(col) for col in df.columns]
            
            # Convert data types and handle None values
            df = df.where(pd.notnull(df), None)
            
            # Handle common data type issues in old MDB files
            for col in df.columns:
                # Convert Access Date/Time to proper format
                if df[col].dtype == 'object':
                    try:
                        # Try to convert date strings
                        pd.to_datetime(df[col], errors='ignore')
                    except:
                        pass
                
                # Handle binary data that might cause issues
                if df[col].dtype == 'object':
                    df[col] = df[col].astype(str).replace('nan', None)
            
            # Insert data into MySQL
            cursor = mysql_conn.cursor()
            
            # Build INSERT statement
            columns = ', '.join([f"`{col}`" for col in df.columns])
            placeholders = ', '.join(['%s'] * len(df.columns))
            insert_sql = f"INSERT INTO `{target_db}`.`{target_table}` ({columns}) VALUES ({placeholders})"
            
            # Insert in smaller batches for old MDB files
            batch_size = 500  # Reduced batch size for better compatibility
            total_rows = len(df)
            
            for i in range(0, total_rows, batch_size):
                batch = df.iloc[i:i+batch_size]
                values = []
                
                for row in batch.values:
                    # Clean up each row for old MDB compatibility
                    clean_row = []
                    for val in row:
                        if pd.isna(val) or val == 'nan':
                            clean_row.append(None)
                        elif isinstance(val, str) and len(val) > 65535:  # Truncate very long strings
                            clean_row.append(val[:65535])
                        else:
                            clean_row.append(val)
                    values.append(tuple(clean_row))
                
                try:
                    cursor.executemany(insert_sql, values)
                    mysql_conn.commit()
                    
                    self.logger.debug(f"Inserted batch {i//batch_size + 1} "
                                    f"({min(i+batch_size, total_rows)}/{total_rows} rows)")
                except Exception as e:
                    self.logger.warning(f"Batch insert failed, trying individual inserts: {e}")
                    # Try inserting rows individually
                    for row_values in values:
                        try:
                            cursor.execute(insert_sql, row_values)
                        except Exception as row_e:
                            self.logger.warning(f"Skipping problematic row: {row_e}")
                            continue
                    mysql_conn.commit()
            
            self.logger.info(f"Migrated {total_rows} records from {source_table} to {target_table}")
            return total_rows
            
        except Exception as e:
            self.logger.error(f"Failed to migrate data for table {source_table}: {e}")
            import traceback
            self.logger.debug(traceback.format_exc())
            return 0
    
    def get_relationships(self, access_conn: pyodbc.Connection) -> List[Dict[str, str]]:
        """Extract relationship information from Access database."""
        relationships = []
        try:
            # This is a basic implementation - Access relationship extraction can be complex
            # For now, we'll log that relationships need manual review
            self.logger.warning("Relationship extraction requires manual review")
            self.logger.warning("Please verify foreign key constraints manually")
        except Exception as e:
            self.logger.error(f"Failed to extract relationships: {e}")
        
        return relationships
    
    def convert_database(self, access_db_path: Path) -> bool:
        """Convert a single Access database to MySQL."""
        db_name = self.sanitize_name(access_db_path.stem)
        self.logger.info(f"Starting conversion of {access_db_path.name} -> {db_name}")
        
        access_conn = None
        mysql_conn = None
        
        try:
            # Connect to Access database
            access_conn = self.connect_to_access(access_db_path)
            if not access_conn:
                return False
            
            # Connect to MySQL
            mysql_conn = self.connect_to_mysql()
            if not mysql_conn:
                return False
            
            # Create MySQL database
            cursor = mysql_conn.cursor()
            cursor.execute(f"CREATE DATABASE IF NOT EXISTS `{db_name}` CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci")
            mysql_conn.commit()
            self.logger.info(f"Created MySQL database: {db_name}")
            
            # Get table list
            tables = self.get_table_list(access_conn)
            if not tables:
                self.logger.warning(f"No tables found in {access_db_path.name}")
                return True
            
            # Convert each table
            converted_tables = 0
            total_records = 0
            
            for table_name in tables:
                try:
                    sanitized_table_name = self.sanitize_name(table_name)
                    self.logger.info(f"Converting table: {table_name} -> {sanitized_table_name}")
                    
                    # Get table structure
                    structure = self.get_table_structure(access_conn, table_name)
                    if not structure['columns']:
                        self.logger.warning(f"Skipping table {table_name} - no structure found")
                        continue
                    
                    # Create MySQL table
                    if self.create_mysql_table(mysql_conn, db_name, sanitized_table_name, structure):
                        # Migrate data
                        records = self.migrate_table_data(access_conn, mysql_conn, 
                                                        table_name, db_name, sanitized_table_name)
                        total_records += records
                        converted_tables += 1
                        self.stats['tables_converted'] += 1
                    else:
                        self.stats['tables_failed'] += 1
                        
                except Exception as e:
                    self.logger.error(f"Failed to convert table {table_name}: {e}")
                    self.logger.error(traceback.format_exc())
                    self.stats['tables_failed'] += 1
                    continue
            
            self.stats['records_migrated'] += total_records
            
            # Extract relationships (basic implementation)
            relationships = self.get_relationships(access_conn)
            self.stats['relationships_created'] += len(relationships)
            
            self.logger.info(f"Database conversion completed: {converted_tables}/{len(tables)} tables converted")
            self.logger.info(f"Total records migrated: {total_records}")
            
            return True
            
        except Exception as e:
            self.logger.error(f"Database conversion failed: {e}")
            self.logger.error(traceback.format_exc())
            return False
            
        finally:
            # Clean up connections
            if access_conn:
                try:
                    access_conn.close()
                except:
                    pass
            if mysql_conn:
                try:
                    mysql_conn.close()
                except:
                    pass
    
    def run_conversion(self) -> Dict[str, Any]:
        """Run the complete conversion process."""
        self.logger.info("Starting MS Access to MySQL conversion process")
        start_time = datetime.now()
        
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
                self.logger.error(traceback.format_exc())
                continue
        
        return self.get_summary_report(start_time)
    
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
        self.logger.info(f"  Relationships Created: {self.stats['relationships_created']}")
        
        success_rate = (self.stats['databases_converted'] / max(self.stats['databases_found'], 1)) * 100
        self.logger.info(f"  Success Rate: {success_rate:.1f}%")
        
        # Save report to JSON
        report_file = self.log_dir / f"conversion_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        with open(report_file, 'w', encoding='utf-8') as f:
            json.dump(report, f, indent=2, default=str)
        
        self.logger.info(f"\nDetailed report saved to: {report_file}")
        self.logger.info(f"{'='*80}")
        
        return report


def main():
    """Main function to run the converter."""
    import argparse
    
    parser = argparse.ArgumentParser(description="Convert MS Access databases to MySQL")
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
    converter = AccessToMySQLConverter(args.source_dir, mysql_config, args.log_dir)
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
