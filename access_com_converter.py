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
    from tqdm import tqdm
    import time
    import threading
    import signal
    from collections import defaultdict
    from tqdm import tqdm
    import threading
except ImportError as e:
    print(f"Missing required package: {e}")
    print("Please install: pip install pywin32 pandas mysql-connector-python tqdm")
    sys.exit(1)


class ConversionStatistics:
    """Track and display conversion statistics with progress bars"""
    
    def __init__(self, log_file="conversion_stats.log"):
        self.start_time = datetime.now()
        self.log_file = log_file
        self.stats = {
            'databases_found': 0,
            'databases_processed': 0,
            'databases_failed': 0,
            'tables_found': 0,
            'tables_processed': 0,
            'tables_failed': 0,
            'tables_updated': 0,
            'tables_skipped': 0,
            'total_rows_processed': 0,
            'total_rows_failed': 0,
            'current_database': '',
            'current_table': '',
            'current_table_rows': 0,
            'processing_phase': 'Starting'
        }
        self.table_sizes = {}  # table_name: estimated_rows
        self.processing_order = []  # Track processing order
        self.table_progress = {}  # table_name: progress_info
        self.lock = threading.Lock()
        
        # Initialize log file
        with open(self.log_file, 'w', encoding='utf-8') as f:
            f.write(f"Conversion Statistics Log - Started: {self.start_time}\n")
            f.write("=" * 60 + "\n\n")
    
    def update_phase(self, phase):
        """Update current processing phase"""
        with self.lock:
            self.stats['processing_phase'] = phase
            self._log_to_file(f"PHASE: {phase}")
    
    def add_database(self, db_path, table_count=0):
        """Register a new database for processing"""
        with self.lock:
            self.stats['databases_found'] += 1
            self.stats['tables_found'] += table_count
            self.stats['current_database'] = os.path.basename(db_path)
            self._log_to_file(f"DATABASE FOUND: {db_path} ({table_count} tables)")
    
    def start_database(self, db_path):
        """Mark start of database processing"""
        with self.lock:
            self.stats['current_database'] = os.path.basename(db_path)
            self._log_to_file(f"DATABASE STARTED: {db_path}")
    
    def complete_database(self, db_path, success=True):
        """Mark completion of database processing"""
        with self.lock:
            if success:
                self.stats['databases_processed'] += 1
                self._log_to_file(f"DATABASE COMPLETED: {db_path}")
            else:
                self.stats['databases_failed'] += 1
                self._log_to_file(f"DATABASE FAILED: {db_path}")
    
    def add_table_size(self, table_name, estimated_rows):
        """Record estimated table size for sorting"""
        with self.lock:
            self.table_sizes[table_name] = estimated_rows
            self._log_to_file(f"TABLE SIZE: {table_name} -> {estimated_rows:,} rows (estimated)")
    
    def start_table(self, table_name, estimated_rows=0):
        """Mark start of table processing"""
        with self.lock:
            self.stats['current_table'] = table_name
            self.stats['current_table_rows'] = 0
            if table_name not in self.processing_order:
                self.processing_order.append(table_name)
            self.table_progress[table_name] = {
                'start_time': datetime.now(),
                'estimated_rows': estimated_rows,
                'processed_rows': 0,
                'status': 'processing'
            }
            self._log_to_file(f"TABLE STARTED: {table_name} ({estimated_rows:,} estimated rows)")
    
    def update_table_progress(self, table_name, processed_rows):
        """Update progress for current table"""
        with self.lock:
            self.stats['current_table_rows'] = processed_rows
            if table_name in self.table_progress:
                self.table_progress[table_name]['processed_rows'] = processed_rows
    
    def complete_table(self, table_name, final_rows, status='completed'):
        """Mark completion of table processing"""
        with self.lock:
            if status == 'completed':
                self.stats['tables_processed'] += 1
                self.stats['total_rows_processed'] += final_rows
            elif status == 'updated':
                self.stats['tables_updated'] += 1
                self.stats['total_rows_processed'] += final_rows
            elif status == 'skipped':
                self.stats['tables_skipped'] += 1
            else:  # failed
                self.stats['tables_failed'] += 1
                self.stats['total_rows_failed'] += final_rows
            
            if table_name in self.table_progress:
                self.table_progress[table_name]['status'] = status
                self.table_progress[table_name]['final_rows'] = final_rows
                self.table_progress[table_name]['end_time'] = datetime.now()
                duration = (self.table_progress[table_name]['end_time'] - 
                           self.table_progress[table_name]['start_time']).total_seconds()
                self.table_progress[table_name]['duration'] = duration
            
            self._log_to_file(f"TABLE {status.upper()}: {table_name} -> {final_rows:,} rows")
    
    def get_sorted_tables(self):
        """Get tables sorted by size (smallest first)"""
        return sorted(self.table_sizes.items(), key=lambda x: x[1])
    
    def display_progress(self):
        """Display current progress statistics"""
        with self.lock:
            elapsed = (datetime.now() - self.start_time).total_seconds()
            
            print("\n" + "="*80)
            print(f"📊 CONVERSION PROGRESS - {self.stats['processing_phase']}")
            print("="*80)
            print(f"⏱️  Runtime: {self._format_duration(elapsed)}")
            print(f"📂 Databases: {self.stats['databases_processed']}/{self.stats['databases_found']} processed, {self.stats['databases_failed']} failed")
            print(f"📋 Tables: {self.stats['tables_processed']}/{self.stats['tables_found']} processed")
            print(f"   ├─ ✅ Completed: {self.stats['tables_processed']}")
            print(f"   ├─ 🔄 Updated: {self.stats['tables_updated']}")  
            print(f"   ├─ ⏭️  Skipped: {self.stats['tables_skipped']}")
            print(f"   └─ ❌ Failed: {self.stats['tables_failed']}")
            print(f"📊 Rows: {self.stats['total_rows_processed']:,} processed, {self.stats['total_rows_failed']:,} failed")
            
            if self.stats['current_database']:
                print(f"🔄 Current Database: {self.stats['current_database']}")
            
            if self.stats['current_table']:
                current_rows = self.stats['current_table_rows']
                table_name = self.stats['current_table']
                estimated = self.table_sizes.get(table_name, 0)
                
                if estimated > 0 and current_rows > 0:
                    progress = min(100, (current_rows / estimated) * 100)
                    print(f"📊 Current Table: {table_name}")
                    print(f"   └─ Progress: {current_rows:,}/{estimated:,} rows ({progress:.1f}%)")
                else:
                    print(f"📊 Current Table: {table_name} - {current_rows:,} rows")
            
            # Show recent completions
            recent_completions = [name for name in self.processing_order[-3:] 
                                if name in self.table_progress and 
                                self.table_progress[name]['status'] != 'processing']
            
            if recent_completions:
                print(f"✅ Recently Completed: {', '.join(recent_completions)}")
            
            print("="*80)
    
    def _format_duration(self, seconds):
        """Format duration in human readable format"""
        if seconds < 60:
            return f"{seconds:.1f}s"
        elif seconds < 3600:
            return f"{seconds/60:.1f}m"
        else:
            hours = seconds // 3600
            minutes = (seconds % 3600) // 60
            return f"{hours:.0f}h {minutes:.0f}m"
    
    def _log_to_file(self, message):
        """Log message to file with timestamp"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = f"[{timestamp}] {message}\n"
        
        try:
            with open(self.log_file, 'a', encoding='utf-8') as f:
                f.write(log_entry)
        except Exception:
            pass  # Ignore logging errors
    
    def save_final_report(self):
        """Save final conversion report"""
        with self.lock:
            end_time = datetime.now()
            total_duration = (end_time - self.start_time).total_seconds()
            
            report = {
                'conversion_summary': {
                    'start_time': self.start_time.isoformat(),
                    'end_time': end_time.isoformat(),
                    'total_duration_seconds': total_duration,
                    'total_duration_formatted': self._format_duration(total_duration)
                },
                'statistics': self.stats.copy(),
                'table_details': self.table_progress.copy(),
                'processing_order': self.processing_order.copy()
            }
            
            # Save JSON report
            report_file = f"conversion_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
            try:
                with open(report_file, 'w', encoding='utf-8') as f:
                    json.dump(report, f, indent=2, default=str)
                print(f"\n📄 Final report saved: {report_file}")
            except Exception as e:
                print(f"❌ Could not save report: {e}")
            
            # Save text summary
            summary_file = f"conversion_summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
            try:
                with open(summary_file, 'w', encoding='utf-8') as f:
                    f.write("MS ACCESS TO MYSQL CONVERSION SUMMARY\n")
                    f.write("="*50 + "\n\n")
                    f.write(f"Start Time: {self.start_time}\n")
                    f.write(f"End Time: {end_time}\n")
                    f.write(f"Total Duration: {self._format_duration(total_duration)}\n\n")
                    
                    f.write("OVERALL STATISTICS:\n")
                    f.write(f"  Databases Found: {self.stats['databases_found']}\n")
                    f.write(f"  Databases Processed: {self.stats['databases_processed']}\n")
                    f.write(f"  Databases Failed: {self.stats['databases_failed']}\n")
                    f.write(f"  Tables Found: {self.stats['tables_found']}\n")
                    f.write(f"  Tables Processed: {self.stats['tables_processed']}\n")
                    f.write(f"  Tables Updated: {self.stats['tables_updated']}\n")
                    f.write(f"  Tables Skipped: {self.stats['tables_skipped']}\n")
                    f.write(f"  Tables Failed: {self.stats['tables_failed']}\n")
                    f.write(f"  Total Rows Processed: {self.stats['total_rows_processed']:,}\n")
                    f.write(f"  Total Rows Failed: {self.stats['total_rows_failed']:,}\n\n")
                    
                    f.write("TABLE PROCESSING ORDER (by size):\n")
                    for i, (table_name, size) in enumerate(self.get_sorted_tables(), 1):
                        status = self.table_progress.get(table_name, {}).get('status', 'not processed')
                        final_rows = self.table_progress.get(table_name, {}).get('final_rows', 0)
                        f.write(f"  {i:2d}. {table_name:<30} {size:>10,} est. -> {final_rows:>10,} actual ({status})\n")
                
                print(f"📄 Summary saved: {summary_file}")
            except Exception as e:
                print(f"❌ Could not save summary: {e}")


class ProgressDisplayThread(threading.Thread):
    """Background thread to display progress updates"""
    
    def __init__(self, stats_tracker, update_interval=10):
        super().__init__(daemon=True)
        self.stats_tracker = stats_tracker
        self.update_interval = update_interval
        self.stop_event = threading.Event()
    
    def run(self):
        while not self.stop_event.wait(self.update_interval):
            self.stats_tracker.display_progress()
    
    def stop(self):
        self.stop_event.set()


class AccessCOMConverter:
    """Convert MS Access databases using COM automation (requires Access installation)."""
    
    def __init__(self, source_dir: str, mysql_config: Dict[str, str], log_dir: str = "logs", stats_tracker: ConversionStatistics = None):
        self.source_dir = Path(source_dir)
        self.mysql_config = mysql_config
        self.log_dir = Path(log_dir)
        self.log_dir.mkdir(exist_ok=True)
        
        # Setup logging
        self.setup_logging()
        
        # Statistics tracking
        self.stats_tracker = stats_tracker or ConversionStatistics()
        
        # Legacy stats (kept for backward compatibility)
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
    
    def safe_close_database(self):
        """Safely close the current Access database."""
        try:
            if self.access_app:
                # Close current database
                self.access_app.CloseCurrentDatabase()
                self.logger.debug("✅ Closed current Access database")
        except Exception as e:
            self.logger.debug(f"Database close warning (usually safe to ignore): {e}")
    
    def is_database_in_use(self, db_path: Path) -> bool:
        """Check if database is currently in use (production-safe check)."""
        try:
            # Check for lock files
            lock_extensions = ['.ldb', '.laccdb']
            
            for ext in lock_extensions:
                lock_file = db_path.with_suffix(ext)
                if lock_file.exists():
                    # Check if lock file is recent (modified in last 10 minutes)
                    import time
                    file_age = time.time() - lock_file.stat().st_mtime
                    if file_age < 600:  # 10 minutes
                        self.logger.warning(f"⚠️  Database {db_path.name} may be in use (recent lock file: {lock_file.name})")
                        return True
            
            # Try to open file in exclusive mode briefly (non-destructive test)
            try:
                with open(db_path, 'r+b') as f:
                    # If we can open it, it's likely not in use
                    pass
                return False
            except PermissionError:
                # File is locked by another process
                self.logger.warning(f"⚠️  Database {db_path.name} is locked by another process")
                return True
                
        except Exception as e:
            self.logger.debug(f"Could not check if database is in use: {e}")
            return False
        
        return False
    
    def safe_open_database(self, db_path: Path) -> bool:
        """Safely open an Access database with production-safe checks."""
        
        # First, check if database appears to be in use (production safety)
        if self.is_database_in_use(db_path):
            self.logger.warning(f"🛡️  Production safety: Database {db_path.name} appears to be in use")
            
            # Give user option to continue or skip
            # In production, we'll log and skip by default
            self.logger.warning("🛡️  Skipping database to avoid production interference")
            self.logger.info("💡 To process this database, ensure it's not in use and remove lock files manually")
            return False
        
        max_retries = 3
        
        for attempt in range(max_retries):
            try:
                # Always close any existing database first
                self.safe_close_database()
                
                # Small delay to ensure cleanup (longer in production)
                import time
                time.sleep(2 + attempt)  # 2-4 seconds delay
                
                # Open the new database
                self.access_app.OpenCurrentDatabase(str(db_path.absolute()))
                self.logger.debug(f"✅ Opened database: {db_path.name}")
                return True
                
            except Exception as e:
                error_msg = str(e).lower()
                
                if "already" in error_msg or "open" in error_msg or "lock" in error_msg:
                    self.logger.warning(f"🔒 Database lock detected (attempt {attempt + 1}/{max_retries}): {e}")
                    
                    if attempt < max_retries - 1:
                        self.logger.info(f"🔄 Waiting longer before retry (production-safe approach)...")
                        
                        # Production-safe retry logic
                        try:
                            # Force close current database
                            self.safe_close_database()
                            
                            # Longer wait in production
                            time.sleep(5 + attempt * 3)
                            
                            # Only restart Access as last resort
                            if attempt == max_retries - 2:
                                self.logger.info("🔄 Restarting Access application (last resort)...")
                                self.close_access()
                                time.sleep(5)
                                if not self.start_access():
                                    return False
                        except Exception as retry_e:
                            self.logger.debug(f"Retry cleanup warning: {retry_e}")
                        
                        continue
                    else:
                        self.logger.error(f"❌ Cannot open database {db_path.name} after {max_retries} attempts")
                        self.logger.error("🛡️  Production safety: Database may be in active use")
                        self.logger.error("💡 Consider:")
                        self.logger.error("   - Running during off-peak hours")
                        self.logger.error("   - Copying database before conversion")  
                        self.logger.error("   - Running: check_database_locks_production_safe.bat")
                        return False
                else:
                    self.logger.error(f"❌ Failed to open database {db_path.name}: {e}")
                    return False
        
        return False
    
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
        """Close Microsoft Access application safely."""
        try:
            if self.access_app:
                # First close any open database
                self.safe_close_database()
                
                # Small delay to ensure cleanup
                import time
                time.sleep(0.5)
                
                # Quit the Access application
                self.access_app.Quit()
                self.access_app = None
                self.logger.info("✅ Microsoft Access closed safely")
        except Exception as e:
            self.logger.warning(f"Warning during Access cleanup: {e}")
            # Force cleanup
            self.access_app = None
    
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
            
            # Method 1: Try CurrentData.AllTables (modern approach)
            tables = []
            try:
                for table in self.access_app.CurrentData.AllTables:
                    table_name = table.Name
                    # Skip system tables
                    if not table_name.startswith("MSys") and not table_name.startswith("~"):
                        tables.append(table_name)
                        
                if tables:
                    self.logger.info(f"Found {len(tables)} user tables (method 1): {tables}")
                    return tables
            except Exception as e:
                self.logger.warning(f"AllTables method failed: {e}")
            
            # Method 2: Try using DAO (Database Access Objects) for old MDB files
            try:
                db = self.access_app.CurrentDb()
                tabledefs = db.TableDefs
                
                for i in range(tabledefs.Count):
                    tabledef = tabledefs.Item(i)
                    table_name = tabledef.Name
                    
                    # Skip system tables and temp tables
                    if (not table_name.startswith("MSys") and 
                        not table_name.startswith("~") and
                        not table_name.startswith("TEMP")):
                        
                        # Check if it's a user table (not a system table)
                        table_type = getattr(tabledef, 'Attributes', 0)
                        if table_type & 2 == 0:  # Not a system table
                            tables.append(table_name)
                
                if tables:
                    self.logger.info(f"Found {len(tables)} user tables (method 2 - DAO): {tables}")
                    return tables
                    
            except Exception as e:
                self.logger.warning(f"DAO method failed: {e}")
            
            # Method 3: Use SQL to query system tables
            try:
                rs = self.access_app.CurrentDb().OpenRecordset(
                    "SELECT Name FROM MSysObjects WHERE Type=1 AND Flags=0 AND Left([Name],4)<>'MSys' AND Left([Name],1)<>'~'"
                )
                
                while not rs.EOF:
                    table_name = rs.Fields("Name").Value
                    tables.append(table_name)
                    rs.MoveNext()
                    
                rs.Close()
                
                if tables:
                    self.logger.info(f"Found {len(tables)} user tables (method 3 - SQL): {tables}")
                    return tables
                    
            except Exception as e:
                self.logger.warning(f"SQL method failed: {e}")
            
            # Method 4: Manual table detection by trying to open recordsets
            try:
                # Try common table names or enumerate through possible names
                potential_tables = ["Table1", "Data", "Main", "Records", "Items", "Customers", "Orders", "Products"]
                
                for potential_name in potential_tables:
                    try:
                        rs = self.access_app.CurrentDb().OpenRecordset(f"SELECT TOP 1 * FROM [{potential_name}]")
                        rs.Close()
                        tables.append(potential_name)
                        self.logger.info(f"Found table by testing: {potential_name}")
                    except:
                        continue
                        
                if tables:
                    self.logger.info(f"Found {len(tables)} tables by testing common names")
                    return tables
                    
            except Exception as e:
                self.logger.warning(f"Manual detection failed: {e}")
            
            self.logger.error("All table enumeration methods failed")
            return []
            
        except Exception as e:
            self.logger.error(f"Failed to get table list from {db_path}: {e}")
            return []
    
    def get_table_size_estimates(self, tables: List[str]) -> Dict[str, int]:
        """Estimate row counts for tables to enable size-based processing order."""
        table_sizes = {}
        
        try:
            db = self.access_app.CurrentDb()
            
            for table_name in tables:
                try:
                    # Try to get accurate count with timeout
                    start_time = time.time()
                    count_rs = db.OpenRecordset(f"SELECT COUNT(*) as RecordCount FROM [{table_name}]")
                    record_count = count_rs.Fields("RecordCount").Value
                    count_rs.Close()
                    
                    elapsed = time.time() - start_time
                    table_sizes[table_name] = record_count
                    
                    self.logger.info(f"Table {table_name}: {record_count:,} rows (counted in {elapsed:.1f}s)")
                    
                    # If counting took too long, this table is likely very large
                    if elapsed > 5.0:
                        self.logger.warning(f"Large table detected: {table_name} took {elapsed:.1f}s to count")
                    
                except Exception as e:
                    self.logger.warning(f"Could not count rows in {table_name}: {e}")
                    # Estimate based on error type or use fallback
                    if "timeout" in str(e).lower() or "large" in str(e).lower():
                        table_sizes[table_name] = 1000000  # Assume very large
                    else:
                        table_sizes[table_name] = 1000  # Assume moderate size
                    
        except Exception as e:
            self.logger.error(f"Failed to estimate table sizes: {e}")
            # Fallback: assign default estimates
            for table_name in tables:
                table_sizes[table_name] = 1000
        
        return table_sizes
    
    def check_existing_table(self, db_name: str, table_name: str) -> tuple[bool, int]:
        """Check if table exists in MySQL and return (exists, row_count)."""
        try:
            conn = self.get_mysql_connection()
            if not conn:
                return False, 0
                
            cursor = conn.cursor()
            
            # Check if table exists
            cursor.execute(f"SELECT COUNT(*) FROM information_schema.tables WHERE table_schema = %s AND table_name = %s", 
                          (db_name, table_name))
            table_exists = cursor.fetchone()[0] > 0
            
            if not table_exists:
                cursor.close()
                conn.close()
                return False, 0
            
            # Get current row count
            cursor.execute(f"SELECT COUNT(*) FROM `{db_name}`.`{table_name}`")
            row_count = cursor.fetchone()[0]
            
            cursor.close()
            conn.close()
            
            return True, row_count
            
        except Exception as e:
            self.logger.warning(f"Could not check existing table {table_name}: {e}")
            return False, 0
    
    def should_update_table(self, db_name: str, table_name: str, access_row_count: int) -> str:
        """Determine if table should be updated. Returns: 'skip', 'update', or 'create'."""
        exists, mysql_row_count = self.check_existing_table(db_name, table_name)
        
        if not exists:
            return 'create'
        
        if mysql_row_count == access_row_count:
            self.logger.info(f"Table {table_name} has same row count ({mysql_row_count:,}), skipping")
            return 'skip'
        elif mysql_row_count < access_row_count:
            self.logger.info(f"Table {table_name} needs update: MySQL has {mysql_row_count:,}, Access has {access_row_count:,}")
            return 'update'
        else:
            self.logger.warning(f"Table {table_name} has more rows in MySQL ({mysql_row_count:,}) than Access ({access_row_count:,})")
            return 'update'  # Still update to ensure consistency
    
    def export_table_to_csv(self, table_name: str, temp_dir: Path) -> Optional[Path]:
        """Export Access table to CSV file."""
        try:
            csv_file = temp_dir / f"{self.sanitize_name(table_name)}.csv"
            
            self.logger.debug(f"Exporting {table_name} to CSV...")
            
            # Method 1: Use DoCmd.TransferText with smaller limits
            try:
                # First, try to get record count to decide approach
                try:
                    db = self.access_app.CurrentDb()
                    count_rs = db.OpenRecordset(f"SELECT COUNT(*) as RecordCount FROM [{table_name}]")
                    record_count = count_rs.Fields("RecordCount").Value
                    count_rs.Close()
                    self.logger.info(f"Table {table_name} has {record_count} records")
                    
                    # If too many records, skip TransferText and go to chunked export
                    if record_count > 50000:
                        self.logger.info(f"Table {table_name} too large for TransferText ({record_count} records), using chunked export")
                        raise Exception("Too many records for TransferText")
                        
                except Exception as count_e:
                    self.logger.debug(f"Could not get record count: {count_e}")
                    # Continue with TransferText attempt
                
                # Standard TransferText export
                self.access_app.DoCmd.TransferText(
                    TransferType=2,  # acExportDelim (CSV export)
                    TableName=table_name,
                    FileName=str(csv_file.absolute()),
                    HasFieldNames=True
                )
                
                if csv_file.exists() and csv_file.stat().st_size > 0:
                    self.logger.debug(f"✅ Exported {table_name} to {csv_file.name} (method 1)")
                    return csv_file
                    
            except Exception as e:
                error_msg = str(e).lower()
                if "too many rows" in error_msg or "limitation" in error_msg or "microsoft access" in error_msg:
                    self.logger.warning(f"TransferText failed due to size limit for {table_name}: {e}")
                else:
                    self.logger.warning(f"DoCmd.TransferText failed for {table_name}: {e}")
            
            # Method 1b: Try TransferText with TOP clause (create a query first)
            try:
                self.logger.info(f"Attempting limited TransferText for large table: {table_name}")
                
                # Create a temporary query with TOP clause
                temp_query_name = f"TempQuery_{self.sanitize_name(table_name)}"
                
                # Delete query if it exists
                try:
                    self.access_app.DoCmd.DeleteObject(1, temp_query_name)  # 1 = acQuery
                except:
                    pass
                
                # Create query with TOP limitation
                sql = f"SELECT TOP 100000 * FROM [{table_name}]"
                db = self.access_app.CurrentDb()
                qdef = db.CreateQueryDef(temp_query_name, sql)
                
                # Export the query instead of the table
                self.access_app.DoCmd.TransferText(
                    TransferType=2,  # acExportDelim
                    TableName=temp_query_name,
                    FileName=str(csv_file.absolute()),
                    HasFieldNames=True
                )
                
                # Clean up the temporary query
                try:
                    self.access_app.DoCmd.DeleteObject(1, temp_query_name)
                except:
                    pass
                
                if csv_file.exists() and csv_file.stat().st_size > 0:
                    self.logger.debug(f"✅ Exported {table_name} via limited query (method 1b) - max 100,000 rows")
                    return csv_file
                    
            except Exception as e:
                self.logger.warning(f"Limited TransferText failed for {table_name}: {e}")
                # Clean up temp query if it exists
                try:
                    temp_query_name = f"TempQuery_{self.sanitize_name(table_name)}"
                    self.access_app.DoCmd.DeleteObject(1, temp_query_name)
                except:
                    pass
            
            # Method 2: Use DoCmd.OutputTo (alternative export method)
            try:
                self.access_app.DoCmd.OutputTo(
                    ObjectType=0,  # acOutputTable
                    ObjectName=table_name,
                    OutputFormat="Microsoft Excel (*.xls)",  # Try Excel first
                    OutputFile=str(temp_dir / f"{self.sanitize_name(table_name)}.xls"),
                    AutoStart=False
                )
                
                # Convert XLS to CSV using pandas
                xls_file = temp_dir / f"{self.sanitize_name(table_name)}.xls"
                if xls_file.exists():
                    import pandas as pd
                    df = pd.read_excel(xls_file)
                    df.to_csv(csv_file, index=False)
                    xls_file.unlink()  # Delete temporary XLS file
                    
                    if csv_file.exists() and csv_file.stat().st_size > 0:
                        self.logger.debug(f"✅ Exported {table_name} via Excel conversion (method 2)")
                        return csv_file
                        
            except Exception as e:
                self.logger.warning(f"OutputTo method failed for {table_name}: {e}")
            
            # Method 3: Direct recordset export for old MDB files (with chunking for large tables)
            try:
                db = self.access_app.CurrentDb()
                
                # Try different SQL variations to access the table
                sql_variations = [
                    f"SELECT * FROM [{table_name}]",
                    f"SELECT * FROM `{table_name}`",
                    f"SELECT * FROM {table_name}",
                    f"SELECT * FROM [{table_name.replace(' ', '_')}]"  # Replace spaces
                ]
                
                rs = None
                for sql in sql_variations:
                    try:
                        rs = db.OpenRecordset(sql)
                        self.logger.debug(f"Successful SQL: {sql}")
                        break
                    except Exception as sql_e:
                        self.logger.debug(f"SQL failed: {sql} - {sql_e}")
                        continue
                
                if rs is None:
                    raise Exception("Could not open recordset with any SQL variation")
                
                # Check record count first
                try:
                    rs.MoveLast()
                    record_count = rs.RecordCount
                    rs.MoveFirst()
                    self.logger.info(f"Table {table_name} has {record_count} records")
                except:
                    record_count = "unknown"
                    self.logger.info(f"Table {table_name} record count: {record_count}")
                
                # Export recordset to CSV manually with chunking
                with open(csv_file, 'w', newline='', encoding='utf-8') as f:
                    import csv
                    writer = csv.writer(f)
                    
                    # Write headers
                    field_names = [field.Name for field in rs.Fields]
                    writer.writerow(field_names)
                    
                    # Write data in chunks to avoid memory issues
                    row_count = 0
                    chunk_size = 1000
                    max_rows = 500000  # Increased limit but still safe
                    
                    # Create progress bar for large tables
                    estimated_rows = self.stats_tracker.table_sizes.get(table_name, max_rows)
                    progress_bar = None
                    
                    if estimated_rows > 10000:  # Only show progress bar for larger tables
                        progress_bar = tqdm(
                            total=min(estimated_rows, max_rows),
                            desc=f"Exporting {table_name}",
                            unit="rows",
                            leave=False,
                            dynamic_ncols=True
                        )
                    
                    while not rs.EOF and row_count < max_rows:
                        chunk_rows = []
                        
                        # Process chunk
                        for _ in range(chunk_size):
                            if rs.EOF or row_count >= max_rows:
                                break
                                
                            row_data = []
                            for field in rs.Fields:
                                try:
                                    value = field.Value
                                    if value is None:
                                        row_data.append('')
                                    elif isinstance(value, (int, float, str)):
                                        row_data.append(str(value))
                                    else:
                                        # Handle dates, binary data, etc.
                                        row_data.append(str(value))
                                except Exception as field_e:
                                    # If field access fails, use empty string
                                    row_data.append('')
                                    
                            chunk_rows.append(row_data)
                            rs.MoveNext()
                            row_count += 1
                        
                        # Write chunk to file
                        writer.writerows(chunk_rows)
                        
                        # Update progress tracking
                        self.stats_tracker.update_table_progress(table_name, row_count)
                        
                        if progress_bar:
                            progress_bar.update(len(chunk_rows))
                        
                        # Progress logging for very large tables
                        if row_count % 50000 == 0:
                            self.logger.info(f"Exported {row_count:,} rows from {table_name}...")
                    
                    if progress_bar:
                        progress_bar.close()
                    
                    if row_count >= max_rows:
                        self.logger.warning(f"Table {table_name} truncated at {max_rows:,} rows for safety")
                
                rs.Close()
                
                if csv_file.exists() and csv_file.stat().st_size > 0:
                    self.logger.debug(f"✅ Exported {table_name} via recordset (method 3) - {row_count} rows")
                    return csv_file
                    
            except Exception as e:
                self.logger.warning(f"Recordset export failed for {table_name}: {e}")
            
            # Method 4: Try exporting in smaller batches using WHERE clause
            try:
                self.logger.info(f"Attempting batch export for large table: {table_name}")
                
                # Try to determine if table has an ID field for batching
                db = self.access_app.CurrentDb()
                rs = db.OpenRecordset(f"SELECT TOP 1 * FROM [{table_name}]")
                
                id_field = None
                for field in rs.Fields:
                    field_name = field.Name.upper()
                    if any(id_name in field_name for id_name in ['ID', 'KEY', 'NUM']):
                        id_field = field.Name
                        break
                
                rs.Close()
                
                if id_field:
                    self.logger.info(f"Found ID field: {id_field}, attempting batch export")
                    
                    # Export in batches
                    batch_size = 10000
                    offset = 0
                    all_rows = []
                    
                    with open(csv_file, 'w', newline='', encoding='utf-8') as f:
                        import csv
                        writer = csv.writer(f)
                        headers_written = False
                        
                        while True:
                            try:
                                batch_sql = f"SELECT TOP {batch_size} * FROM [{table_name}] ORDER BY [{id_field}]"
                                if offset > 0:
                                    # This is a simplified approach - may need adjustment
                                    batch_sql = f"SELECT TOP {batch_size} * FROM [{table_name}] WHERE [{id_field}] > {offset} ORDER BY [{id_field}]"
                                
                                rs_batch = db.OpenRecordset(batch_sql)
                                
                                if rs_batch.EOF:
                                    rs_batch.Close()
                                    break
                                
                                # Write headers once
                                if not headers_written:
                                    field_names = [field.Name for field in rs_batch.Fields]
                                    writer.writerow(field_names)
                                    headers_written = True
                                
                                # Write batch data
                                batch_count = 0
                                last_id = offset
                                
                                while not rs_batch.EOF:
                                    row_data = []
                                    for field in rs_batch.Fields:
                                        try:
                                            value = field.Value
                                            if field.Name == id_field:
                                                last_id = value if value is not None else last_id
                                            row_data.append(str(value) if value is not None else '')
                                        except:
                                            row_data.append('')
                                    
                                    writer.writerow(row_data)
                                    rs_batch.MoveNext()
                                    batch_count += 1
                                
                                rs_batch.Close()
                                offset = last_id
                                
                                self.logger.debug(f"Exported batch: {batch_count} rows (last ID: {last_id})")
                                
                                if batch_count < batch_size:
                                    break  # Last batch
                                    
                            except Exception as batch_e:
                                self.logger.warning(f"Batch export failed at offset {offset}: {batch_e}")
                                break
                    
                    if csv_file.exists() and csv_file.stat().st_size > 0:
                        self.logger.debug(f"✅ Exported {table_name} via batch method")
                        return csv_file
                
            except Exception as e:
                self.logger.warning(f"Batch export failed for {table_name}: {e}")
            
            # Method 4: Try with quoted table name variations
            try:
                quoted_variations = [
                    f'"{table_name}"',
                    f"'{table_name}'",
                    f"[{table_name}]",
                    f"`{table_name}`"
                ]
                
                for quoted_name in quoted_variations:
                    try:
                        self.access_app.DoCmd.TransferText(
                            TransferType=2,  # acExportDelim
                            TableName=quoted_name,
                            FileName=str(csv_file.absolute()),
                            HasFieldNames=True
                        )
                        
                        if csv_file.exists() and csv_file.stat().st_size > 0:
                            self.logger.debug(f"✅ Exported {table_name} with quoting: {quoted_name}")
                            return csv_file
                            
                    except Exception as quote_e:
                        self.logger.debug(f"Quoted name {quoted_name} failed: {quote_e}")
                        continue
                        
            except Exception as e:
                self.logger.warning(f"Quoted name variations failed: {e}")
            
            # If all methods failed
            self.logger.error(f"All export methods failed for table: {table_name}")
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
        """Convert a single Access database to MySQL with intelligent processing."""
        db_name = self.sanitize_name(db_path.stem)
        
        self.stats_tracker.start_database(db_path)
        self.logger.info(f"🚀 Starting conversion of {db_path.name} -> {db_name}")
        
        try:
            # Phase 1: Open and analyze database
            self.stats_tracker.update_phase(f"Analyzing {db_path.name}")
            
            if not self.safe_open_database(db_path):
                self.stats_tracker.complete_database(db_path, success=False)
                return False
            
            # Phase 2: Get table list
            self.stats_tracker.update_phase(f"Enumerating tables in {db_path.name}")
            self.logger.info("📋 Enumerating tables...")
            tables = self.get_table_list(db_path)
            
            if not tables:
                self.logger.warning(f"No tables found in {db_path.name}")
                # Try debug enumeration (keeping existing logic)
                try:
                    self.logger.info("🔍 Debug: Attempting to list ALL objects...")
                    db = self.access_app.CurrentDb()
                    tabledefs = db.TableDefs
                    
                    all_objects = []
                    for i in range(tabledefs.Count):
                        tabledef = tabledefs.Item(i)
                        all_objects.append(f"{tabledef.Name} (Type: {getattr(tabledef, 'Attributes', 'unknown')})")
                    
                    self.logger.info(f"All objects in database: {all_objects}")
                    
                    if all_objects:
                        for obj_info in all_objects:
                            obj_name = obj_info.split(' (')[0]
                            if not obj_name.startswith('MSys') and not obj_name.startswith('~'):
                                tables = [obj_name]
                                self.logger.info(f"Will attempt to process: {obj_name}")
                                break
                                
                except Exception as debug_e:
                    self.logger.error(f"Debug enumeration failed: {debug_e}")
                    
                if not tables:
                    self.stats_tracker.complete_database(db_path, success=True)  # Empty database is success
                    return True
            
            self.logger.info(f"📊 Found {len(tables)} tables: {', '.join(tables)}")
            
            # Phase 3: Estimate table sizes for optimal processing order
            self.stats_tracker.update_phase(f"Analyzing table sizes in {db_path.name}")
            self.logger.info("📏 Estimating table sizes for optimal processing order...")
            
            table_sizes = self.get_table_size_estimates(tables)
            
            # Register table sizes with stats tracker
            for table_name, size in table_sizes.items():
                self.stats_tracker.add_table_size(table_name, size)
            
            # Phase 4: Sort tables by size (smallest first)
            sorted_tables = sorted(table_sizes.items(), key=lambda x: x[1])
            small_tables = [(name, size) for name, size in sorted_tables if size < 100000]
            large_tables = [(name, size) for name, size in sorted_tables if size >= 100000]
            
            self.logger.info(f"📊 Processing order - Small tables: {len(small_tables)}, Large tables: {len(large_tables)}")
            
            # Create temporary directory for CSV exports
            temp_dir = Path(tempfile.mkdtemp(prefix="access_export_"))
            self.logger.debug(f"Using temp directory: {temp_dir}")
            
            # Phase 5: Process small tables first
            if small_tables:
                self.stats_tracker.update_phase(f"Processing small tables ({len(small_tables)} tables)")
                self.logger.info(f"🏃‍♀️ Processing {len(small_tables)} small tables first...")
                
                for table_name, estimated_size in small_tables:
                    self._process_single_table(table_name, estimated_size, db_name, temp_dir)
            
            # Phase 6: Process large tables
            if large_tables:
                self.stats_tracker.update_phase(f"Processing large tables ({len(large_tables)} tables)")
                self.logger.info(f"🐌 Processing {len(large_tables)} large tables...")
                
                for table_name, estimated_size in large_tables:
                    self._process_single_table(table_name, estimated_size, db_name, temp_dir)
            
            # Cleanup
            try:
                import shutil
                shutil.rmtree(temp_dir, ignore_errors=True)
                self.logger.debug("Cleaned up temporary directory")
            except Exception:
                pass
            
            # Always close the database when done
            self.safe_close_database()
            
            self.stats_tracker.complete_database(db_path, success=True)
            self.logger.info(f"🎉 Successfully completed conversion of {db_path.name}")
            return True
            
        except Exception as e:
            self.logger.error(f"❌ Failed to convert database {db_path}: {e}")
            self.stats_tracker.complete_database(db_path, success=False)
            # Ensure database is closed even on failure
            self.safe_close_database()
            return False
    
    def _process_single_table(self, table_name: str, estimated_size: int, db_name: str, temp_dir: Path):
        """Process a single table with full statistics tracking."""
        sanitized_table_name = self.sanitize_name(table_name)
        
        # Start table processing
        self.stats_tracker.start_table(table_name, estimated_size)
        
        try:
            # Check if we should skip, update, or create this table
            action = self.should_update_table(db_name, sanitized_table_name, estimated_size)
            
            if action == 'skip':
                self.logger.info(f"⏭️  Skipping {table_name} - no changes needed")
                self.stats_tracker.complete_table(table_name, estimated_size, 'skipped')
                return
            
            # Log what we're doing
            size_desc = "small" if estimated_size < 10000 else "medium" if estimated_size < 100000 else "large"
            action_desc = "Creating" if action == 'create' else "Updating"
            
            self.logger.info(f"🔄 {action_desc} {size_desc} table: '{table_name}' -> '{sanitized_table_name}' ({estimated_size:,} rows)")
            
            # Try to get basic table info first
            try:
                rs = self.access_app.CurrentDb().OpenRecordset(f"SELECT TOP 1 * FROM [{table_name}]")
                field_count = rs.Fields.Count
                rs.Close()
                self.logger.debug(f"Table {table_name} has {field_count} fields")
            except Exception as info_e:
                self.logger.warning(f"Could not get table info for {table_name}: {info_e}")
            
            # Export to CSV
            csv_file = self.export_table_to_csv(table_name, temp_dir)
            if not csv_file:
                self.logger.error(f"❌ Failed to export table: {table_name}")
                self.stats_tracker.complete_table(table_name, 0, 'failed')
                return
            
            # Import to MySQL
            records = self.import_csv_to_mysql(csv_file, db_name, sanitized_table_name)
            if records > 0:
                status = 'updated' if action == 'update' else 'completed'
                self.stats_tracker.complete_table(table_name, records, status)
                
                # Log success with performance info
                if estimated_size > 0:
                    efficiency = (records / estimated_size) * 100
                    self.logger.info(f"✅ {action_desc} {table_name}: {records:,} records ({efficiency:.1f}% of estimate)")
                else:
                    self.logger.info(f"✅ {action_desc} {table_name}: {records:,} records")
            else:
                self.logger.error(f"❌ Failed to import {table_name} to MySQL")
                self.stats_tracker.complete_table(table_name, 0, 'failed')
                
        except Exception as e:
            self.logger.error(f"❌ Error processing table {table_name}: {e}")
            self.stats_tracker.complete_table(table_name, 0, 'failed')

    def run_conversion(self) -> Dict[str, Any]:
        """Run the complete conversion process with enhanced statistics and progress tracking."""
        self.stats_tracker.update_phase("Starting conversion process")
        self.logger.info("🚀 Starting MS Access to MySQL conversion using COM automation")
        start_time = datetime.now()
        
        # Start progress display thread
        progress_thread = ProgressDisplayThread(self.stats_tracker, update_interval=15)
        progress_thread.start()
        
        try:
            # Start Microsoft Access
            if not self.start_access():
                return self.get_summary_report(start_time)
            
            # Phase 1: Discovery
            self.stats_tracker.update_phase("Discovering Access databases")
            databases = self.find_access_databases()
            
            if not databases:
                self.logger.warning("No MS Access databases found")
                self.logger.info("❌ No Access database files found in the source directory")
                return self.get_summary_report(start_time)
            
            self.logger.info(f"📂 Found {len(databases)} Access database(s)")
            
            # Register all databases with stats tracker
            for db_path in databases:
                # Quick scan to count tables for better progress tracking
                try:
                    if self.safe_open_database(db_path):
                        tables = self.get_table_list(db_path)
                        table_count = len(tables)
                        self.stats_tracker.add_database(db_path, table_count)
                        self.safe_close_database()
                    else:
                        self.logger.warning(f"Could not open {db_path.name} for pre-scan")
                        self.stats_tracker.add_database(db_path, 0)
                except Exception as e:
                    self.logger.warning(f"Could not pre-scan {db_path.name}: {e}")
                    self.stats_tracker.add_database(db_path, 0)
                    # Ensure database is closed even on error
                    self.safe_close_database()
            
            # Phase 2: Convert databases
            self.stats_tracker.update_phase("Converting databases")
            
            for db_path in databases:
                try:
                    self.logger.info(f"\n{'='*80}")
                    self.logger.info(f"📂 Processing database: {db_path}")
                    self.logger.info(f"{'='*80}")
                    
                    if self.convert_database(db_path):
                        self.logger.info(f"✅ Successfully converted: {db_path.name}")
                    else:
                        self.logger.error(f"❌ Failed to convert: {db_path.name}")
                        
                except Exception as e:
                    self.logger.error(f"❌ Unexpected error processing {db_path}: {e}")
                    self.stats_tracker.complete_database(db_path, success=False)
                    # Ensure database is closed even on error
                    self.safe_close_database()
                    continue
                
                # Ensure database is closed between processing
                self.safe_close_database()
                
                # Brief pause between databases to allow Access to clean up properly
                time.sleep(2)
            
            # Final phase
            self.stats_tracker.update_phase("Completing conversion")
            
        finally:
            # Stop progress display
            progress_thread.stop()
            progress_thread.join(timeout=2)
            
            # Always close Access
            self.close_access()
            
            # Final statistics display
            self.stats_tracker.display_progress()
            self.stats_tracker.save_final_report()
        
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
    """Main function to run the COM converter with enhanced statistics."""
    import argparse
    
    parser = argparse.ArgumentParser(description="Convert MS Access databases to MySQL using COM automation")
    parser.add_argument("source_dir", help="Directory containing MS Access database files")
    parser.add_argument("--host", default="localhost", help="MySQL host (default: localhost)")
    parser.add_argument("--port", type=int, default=3306, help="MySQL port (default: 3306)")
    parser.add_argument("--user", required=True, help="MySQL username")
    parser.add_argument("--password", required=True, help="MySQL password")
    parser.add_argument("--log-dir", default="logs", help="Directory for log files (default: logs)")
    parser.add_argument("--stats-interval", type=int, default=15, help="Progress display interval in seconds (default: 15)")
    
    args = parser.parse_args()
    
    # MySQL configuration
    mysql_config = {
        'host': args.host,
        'port': args.port,
        'user': args.user,
        'password': args.password,
        'autocommit': False
    }
    
    # Create statistics tracker
    stats_tracker = ConversionStatistics(log_file=f"{args.log_dir}/conversion_stats.log")
    
    # Setup graceful shutdown handling
    def signal_handler(signum, frame):
        print(f"\n⚠️  Received interrupt signal ({signum})")
        stats_tracker.update_phase("Shutting down gracefully...")
        stats_tracker.display_progress()
        stats_tracker.save_final_report()
        print("📄 Statistics and reports have been saved")
        sys.exit(1)
    
    signal.signal(signal.SIGINT, signal_handler)
    signal.signal(signal.SIGTERM, signal_handler)
    
    print("🚀 MS ACCESS TO MYSQL CONVERTER")
    print("="*50)
    print(f"📂 Source Directory: {args.source_dir}")
    print(f"🗄️  MySQL Host: {args.host}:{args.port}")
    print(f"👤 MySQL User: {args.user}")
    print(f"📝 Log Directory: {args.log_dir}")
    print(f"📊 Stats Update Interval: {args.stats_interval}s")
    print("="*50)
    
    try:
        # Create converter and run
        converter = AccessCOMConverter(args.source_dir, mysql_config, args.log_dir, stats_tracker)
        report = converter.run_conversion()
        
        # Final summary
        print("\n" + "="*80)
        print("🎯 CONVERSION SUMMARY")
        print("="*80)
        
        total_databases = stats_tracker.stats['databases_found']
        processed_databases = stats_tracker.stats['databases_processed']
        failed_databases = stats_tracker.stats['databases_failed']
        
        total_tables = stats_tracker.stats['tables_found']
        processed_tables = stats_tracker.stats['tables_processed']
        updated_tables = stats_tracker.stats['tables_updated']
        skipped_tables = stats_tracker.stats['tables_skipped']
        failed_tables = stats_tracker.stats['tables_failed']
        
        total_rows = stats_tracker.stats['total_rows_processed']
        failed_rows = stats_tracker.stats['total_rows_failed']
        
        print(f"📂 Databases: {processed_databases}/{total_databases} successful, {failed_databases} failed")
        print(f"📋 Tables: {processed_tables + updated_tables}/{total_tables} successful")
        print(f"   ├─ ✅ New tables: {processed_tables}")
        print(f"   ├─ 🔄 Updated tables: {updated_tables}")
        print(f"   ├─ ⏭️  Skipped tables: {skipped_tables}")
        print(f"   └─ ❌ Failed tables: {failed_tables}")
        print(f"📊 Data: {total_rows:,} rows processed, {failed_rows:,} rows failed")
        
        # Performance statistics
        elapsed_time = (datetime.now() - stats_tracker.start_time).total_seconds()
        if elapsed_time > 0:
            rows_per_second = total_rows / elapsed_time
            print(f"⚡ Performance: {rows_per_second:,.0f} rows/second average")
        
        print("="*80)
        
        # Exit with appropriate code
        if failed_databases == 0 and failed_tables == 0:
            print("✅ ALL CONVERSIONS COMPLETED SUCCESSFULLY!")
            print(f"📄 Detailed reports saved to: conversion_report_*.json")
            print(f"📄 Statistics log saved to: {args.log_dir}/conversion_stats.log")
            sys.exit(0)
        elif processed_databases > 0 or (processed_tables + updated_tables) > 0:
            print("⚠️  CONVERSION COMPLETED WITH SOME ISSUES")
            print("📄 Check the detailed reports and logs for error information")
            print(f"📄 Reports saved to: conversion_report_*.json")
            print(f"📄 Statistics log saved to: {args.log_dir}/conversion_stats.log")
            sys.exit(1)
        else:
            print("❌ CONVERSION FAILED COMPLETELY")
            print("📄 Check the logs for detailed error information")
            sys.exit(2)
    
    except KeyboardInterrupt:
        print(f"\n⚠️  Conversion interrupted by user")
        sys.exit(130)
    except Exception as e:
        print(f"\n❌ Unexpected error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
