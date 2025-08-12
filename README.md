# MS Access to MySQL Database Converter

A comprehensive Python tool for automatically converting Microsoft Access databases (.mdb, .accdb) to MySQL databases with full structure preservation, data migration, and relationship handling.

## Features

- **Automatic Database Discovery**: Recursively finds all MS Access databases in a directory
- **Complete Structure Migration**: Converts tables, columns, data types, and constraints
- **Data Migration**: Transfers all records with proper type conversion
- **Relationship Preservation**: Attempts to maintain foreign key relationships
- **Error Resilience**: Continues processing even when individual databases fail
- **Comprehensive Logging**: Detailed logs with progress tracking and error reporting
- **Generic Operation**: Works with any Access database structure without prior knowledge
- **Batch Processing**: Handles multiple databases in a single run
- **Progress Tracking**: Real-time status updates and statistics
- **Configuration Management**: Save and reuse connection settings

## Requirements

- Python 3.7 or higher
- Microsoft Access Database Engine (for reading .mdb/.accdb files)
- MySQL Server (target database)

### Python Dependencies

```bash
pip install -r requirements.txt
```

Required packages:
- `pyodbc` - For connecting to MS Access databases
- `mysql-connector-python` - For connecting to MySQL
- `pandas` - For data manipulation and transfer

### System Requirements

**Windows:**
- Microsoft Access Database Engine 2016 Redistributable
- Or Microsoft Office with Access installed

**Alternative for Linux/Mac:**
- Use mdb-tools or similar utilities (additional configuration required)

## Quick Start

### For Your Situation (Access 16.0 installed, old MDB files)

Since you have Microsoft Access 16.0 installed, use the **COM automation converter** which is perfect for old MDB files:

```bash
# Direct conversion using your installed Access
python access_com_converter.py "C:\path\to\your\mdb\files" --user your_mysql_user --password your_mysql_password

# Example:
python access_com_converter.py "C:\databases" --user root --password mypassword --host localhost
```

### Alternative Methods

#### **Option 1: ODBC Method (requires driver installation)**

#### **Option 1: ODBC Method (requires driver installation)**

First install Microsoft Access Database Engine, then:

```bash
# Install dependencies
install.bat

# Configure settings
python config_setup.py setup

# Run conversion
run_conversion.bat
```

#### **Option 2: Manual CSV Export**

If automation fails:

```bash
# 1. Export tables from Access to CSV manually
# 2. Use CSV converter
python csv_to_mysql_converter.py "C:\path\to\csv\files" --user root --password mypass
```

## Detailed Usage

### Using COM Automation (Recommended for your setup)

### Using COM Automation (Recommended for your setup)

The COM automation method uses your installed Microsoft Access to export tables and convert them to MySQL. This is the most reliable method for old MDB files.

```bash
# Basic usage - convert all MDB/ACCDB files in a directory
python access_com_converter.py "C:\path\to\databases" --user root --password secret

# With custom MySQL server
python access_com_converter.py "C:\databases" --host 192.168.1.100 --user dbuser --password dbpass

# With custom log directory
python access_com_converter.py "C:\databases" --user root --password secret --log-dir "C:\conversion_logs"
```

**Advantages of COM method:**
- ✅ Works with ANY Access database version (including very old MDB files)
- ✅ No ODBC driver installation required
- ✅ Handles complex Access features better
- ✅ More reliable for old/corrupted databases
- ✅ Uses your existing Access installation

### Using ODBC Method

### Using Configuration File

```bash
# Create configuration
python config_setup.py setup

# Run with saved configuration
python run_converter.py

# Use custom configuration file
python run_converter.py --config my_config.json
```

## File Structure

```
msaccess-script/
├── access_com_converter.py         # COM automation converter (RECOMMENDED for your setup)
├── access_to_mysql_converter.py    # Main ODBC-based conversion engine
├── legacy_mdb_converter.py         # Multi-method converter for old MDB files
├── csv_to_mysql_converter.py       # Manual CSV import tool
├── config_setup.py                 # Interactive configuration setup
├── run_converter.py                # Convenient runner script
├── diagnose_odbc.py                # ODBC diagnostics and troubleshooting
├── fix_odbc_drivers.bat            # Windows ODBC driver fix script
├── example_usage.py                # Sample programmatic usage
├── requirements.txt                # Python dependencies
├── README.md                       # This file
├── install.bat                     # Windows installation script
├── run_conversion.bat              # Windows quick start script
├── logs/                           # Generated log files
│   ├── access_com_converter_YYYYMMDD_HHMMSS.log
│   ├── access_to_mysql_YYYYMMDD_HHMMSS.log
│   └── conversion_report_YYYYMMDD_HHMMSS.json
└── converter_config.json          # Saved configuration (generated)
```

## Configuration Options

### MySQL Connection
- **Host**: MySQL server hostname or IP
- **Port**: MySQL server port (default: 3306)
- **User**: MySQL username
- **Password**: MySQL password

### Advanced Options
- **Batch Size**: Number of records to process at once (default: 1000)
- **Include System Tables**: Whether to convert Access system tables
- **Create Indexes**: Auto-create indexes based on Access indexes
- **Character Encoding**: MySQL character encoding (default: utf8mb4)

## Data Type Mapping

The converter automatically maps Access data types to MySQL equivalents:

| Access Type | MySQL Type |
|-------------|------------|
| COUNTER | INT AUTO_INCREMENT PRIMARY KEY |
| LONG | INT |
| INTEGER | INT |
| SHORT | SMALLINT |
| BYTE | TINYINT |
| SINGLE | FLOAT |
| DOUBLE | DOUBLE |
| CURRENCY | DECIMAL(19,4) |
| DATETIME | DATETIME |
| BIT | BOOLEAN |
| TEXT | VARCHAR(size) or TEXT |
| MEMO | TEXT |
| LONGBINARY | LONGBLOB |
| BINARY | VARBINARY(255) |

## Logging and Monitoring

### Log Files

All operations are logged with timestamps and detailed information:

```
logs/
├── access_to_mysql_YYYYMMDD_HHMMSS.log  # Detailed execution log
└── conversion_report_YYYYMMDD_HHMMSS.json  # Summary report
```

### Log Contents

- Database discovery and processing
- Table structure conversion
- Data migration progress
- Error messages and stack traces
- Performance statistics
- Summary reports

### Monitoring Progress

The script provides real-time feedback:
- Database discovery results
- Individual table conversion status
- Record migration counts
- Error notifications
- Final summary with success rates

## Error Handling

The converter is designed to be resilient:

- **Database-level errors**: Skips failed databases, continues with others
- **Table-level errors**: Skips failed tables, continues with other tables in the database
- **Data-level errors**: Logs problematic records but continues migration
- **Connection errors**: Attempts to reconnect and provides clear error messages

## Troubleshooting

### Common Issues

1. **"Data source name not found and no default driver specified" (ODBC Error)**
   
   **Cause**: Microsoft Access ODBC driver is not installed or not compatible with your Python architecture.
   
   **For OLD .MDB FILES (your case):**
   ```bash
   # Quick fix script
   fix_odbc_drivers.bat
   
   # Or use the legacy converter
   python legacy_mdb_converter.py "C:\path\to\old\mdb\files" --user myuser --password mypass
   ```
   
   **Manual Fix for Old MDB Files**:
   - Download Microsoft Access Database Engine 2016 Redistributable:
     https://www.microsoft.com/en-us/download/details.aspx?id=54920
   - **For 64-bit Python**: Install AccessDatabaseEngine_X64.exe
   - **For 32-bit Python**: Install AccessDatabaseEngine.exe
   - If you get "Another version is already installed" error:
     - Run as Administrator: `AccessDatabaseEngine_X64.exe /quiet` (or AccessDatabaseEngine.exe /quiet)
   
   **Alternative for Very Old MDB Files**:
   - Download legacy Jet Database Engine 4.0:
     https://www.microsoft.com/en-us/download/details.aspx?id=23734
   - Use the `legacy_mdb_converter.py` script which tries multiple methods

2. **"Microsoft Access Driver not found"**
   - Install Microsoft Access Database Engine 2016 Redistributable
   - Ensure 32-bit Python with 32-bit driver or 64-bit Python with 64-bit driver

3. **"Permission denied" errors**
   - Ensure Access databases are not open in Microsoft Access
   - Check file permissions
   - Run with administrator privileges if needed

4. **MySQL connection errors**
   - Verify MySQL server is running
   - Check firewall settings
   - Confirm username/password and permissions

5. **Memory errors with large databases**
   - Reduce batch size in configuration
   - Ensure sufficient system RAM
   - Close other applications

### Debug Mode

For additional debugging information, check the detailed log files in the `logs` directory.

## Limitations

- **Relationships**: Complex Access relationships may require manual review
- **Queries**: Access queries are not converted (tables and data only)
- **Forms/Reports**: Only table structure and data are migrated
- **Custom Functions**: Access-specific functions are not converted
- **Security**: Access user-level security is not migrated

## Performance Tips

- **Batch Size**: Adjust based on available memory and database size
- **Network**: Use local MySQL server when possible for faster data transfer
- **Disk Space**: Ensure adequate space for MySQL data files
- **Indexes**: Let the converter create indexes, then optimize as needed

## Security Considerations

- Configuration files may contain passwords - protect appropriately
- Use MySQL users with appropriate permissions only
- Consider using MySQL SSL connections for remote servers
- Back up original Access databases before conversion

## Contributing

To contribute to this project:

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests if applicable
5. Submit a pull request

## License

This project is open source. Please check the LICENSE file for details.

## Support

For issues, questions, or feature requests:

1. Check the troubleshooting section above
2. Review log files for detailed error information
3. Create an issue with:
   - Operating system and Python version
   - Access and MySQL versions
   - Error messages from log files
   - Sample database structure (if possible)

## Version History

- **v1.0**: Initial release with basic conversion functionality
- **v1.1**: Added configuration management and improved error handling
- **v1.2**: Enhanced relationship detection and logging system
