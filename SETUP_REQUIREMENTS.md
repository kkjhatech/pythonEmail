# Setup Requirements and Configuration Guide

## System Requirements

### Minimum System Specifications
- **Operating System**: Windows 10/11, Windows Server 2016+, Linux (Ubuntu 18.04+)
- **Python**: Version 3.8 or higher
- **Memory**: 4GB RAM minimum (8GB recommended)
- **Storage**: 10GB free disk space
- **Network**: Stable internet connection for email access

### Software Dependencies

#### Python Packages (requirements.txt)
```
pyodbc>=4.0.32          # SQL Server database connectivity
pandas>=1.5.0           # Data manipulation and Excel processing
openpyxl>=3.0.0         # Excel .xlsx file support
xlrd>=2.0.0             # Excel .xls file support
python-dotenv>=0.19.0   # Environment variable management
schedule>=1.2.0         # Task scheduling
```

#### Database Requirements
- **SQL Server**: Version 2016 or higher
- **Authentication**: SQL Server Authentication or Windows Authentication
- **Permissions**: CREATE TABLE, INSERT, SELECT, UPDATE, DELETE
- **ODBC Driver**: Microsoft ODBC Driver for SQL Server (version 17+)

#### Email Service Requirements
- **Gmail**: IMAP access enabled, App Password for 2FA
- **Outlook/Office365**: IMAP access enabled or Outlook desktop application
- **Exchange Server**: IMAP access configured

## Installation Steps

### 1. Python Environment Setup
```bash
# Check Python version
python --version

# Create virtual environment
python -m venv email_reader_env

# Activate virtual environment
# Windows
email_reader_env\Scripts\activate
# Linux/Mac
source email_reader_env/bin/activate

# Install dependencies
pip install -r requirements.txt
```

### 2. Database Setup

#### SQL Server Configuration
```sql
-- Create database
CREATE DATABASE EmailAutomation;

-- Create login (if using SQL Server Authentication)
CREATE LOGIN email_reader_user WITH PASSWORD = 'StrongPassword123!';

-- Create user and grant permissions
USE EmailAutomation;
CREATE USER email_reader_user FOR LOGIN email_reader_user;
ALTER ROLE db_owner ADD MEMBER email_reader_user;
```

#### Table Creation Script
```sql
USE EmailAutomation;

-- Email_Master table
CREATE TABLE Email_Master (
    Email_Master_A INT IDENTITY(1,1) PRIMARY KEY,
    EmailID NVARCHAR(255) NOT NULL UNIQUE,
    CreatedDate DATETIME DEFAULT GETDATE(),
    CreatedBy NVARCHAR(100) DEFAULT 'System'
);

-- Email_Details table
CREATE TABLE Email_Details (
    Email_Details_A INT IDENTITY(1,1) PRIMARY KEY,
    EmailID_N INT NOT NULL,
    Subject_Name NVARCHAR(50),
    SheetName NVARCHAR(50) DEFAULT 'Sheet1',
    TotalRows INT,
    ReceivedDate DATETIME,
    FOREIGN KEY (EmailID_N) REFERENCES Email_Master(Email_Master_A)
);

-- Create indexes for performance
CREATE INDEX IX_Email_Master_EmailID ON Email_Master(EmailID);
CREATE INDEX IX_Email_Details_EmailID_N ON Email_Details(EmailID_N);
```

#### Stored Procedures
```sql
-- usp_insert_email procedure
CREATE OR ALTER PROCEDURE [dbo].[usp_insert_email]
    @Email_ID nvarchar(100),
    @CreatedBy nvarchar(100) = 'System'
AS
BEGIN
    SET NOCOUNT ON;
    
    -- Check if email already exists
    IF NOT EXISTS (SELECT 1 FROM Email_Master WHERE EmailID = @Email_ID)
    BEGIN
        INSERT INTO Email_Master (EmailID, CreatedDate, CreatedBy) 
        VALUES (@Email_ID, GETDATE(), @CreatedBy)
    END
END;

-- usp_insert_email_details procedure
CREATE OR ALTER PROCEDURE [dbo].[usp_insert_email_details]
    @EmailID_N int,
    @Subject_Name nvarchar(50),
    @SheetName nvarchar(50),
    @TotalRows int,
    @ReceivedDate datetime
AS
BEGIN
    SET NOCOUNT ON;
    
    INSERT INTO Email_Details (EmailID_N, Subject_Name, SheetName, TotalRows, ReceivedDate)
    VALUES (@EmailID_N, @Subject_Name, @SheetName, @TotalRows, @ReceivedDate)
END;
```

### 3. Email Service Configuration

#### Gmail Setup
1. Enable IMAP in Gmail settings
2. Generate App Password (if 2FA enabled):
   - Go to Google Account settings
   - Security → 2-Step Verification → App passwords
   - Generate new app password for "Mail"

#### Outlook/Office365 Setup
1. For IMAP access:
   - Enable IMAP in Outlook settings
   - Use app password if 2FA enabled

2. For Outlook COM automation:
   - Install Outlook desktop application
   - Ensure Outlook is configured with email account

### 4. Configuration File Setup

#### .env File Template
```env
# =============================================================================
# Database Configuration
# =============================================================================
DB_SERVER=localhost
DB_DATABASE=EmailAutomation
DB_USERNAME=email_reader_user
DB_PASSWORD=StrongPassword123!

# =============================================================================
# Email Configuration
# =============================================================================
EMAIL_PROVIDER=outlook          # Options: gmail, outlook, office365
EMAIL_SERVER=outlook.office365.com
EMAIL_PORT=993
EMAIL_USE_SSL=true
EMAIL_USERNAME=your_email@domain.com
EMAIL_PASSWORD=your_app_password

# =============================================================================
# Outlook COM Configuration (Alternative to IMAP)
# =============================================================================
OUTLOOK_CONNECTION_METHOD=com   # Options: com, imap

# =============================================================================
# Processing Configuration
# =============================================================================
DOWNLOAD_FOLDER=./downloads
LOG_FOLDER=./logs
CHECK_INTERVAL_MINUTES=30

# =============================================================================
# Email Processing Configuration
# =============================================================================
EMAIL_PROCESSED_FOLDER=Processed
EMAIL_MOVE_PROCESSED=true
EMAIL_MARK_READ=true

# =============================================================================
# File Processing Configuration
# =============================================================================
MAX_FILE_SIZE_MB=50
SUPPORTED_EXTENSIONS=.xlsx,.xls,.xlsb,.csv
DELETE_AFTER_PROCESSING=false
```

#### Directory Structure Setup
```bash
# Create required directories
mkdir downloads
mkdir logs
mkdir backups

# Set permissions (Linux/Mac)
chmod 755 downloads logs backups
```

### 5. ODBC Driver Installation

#### Windows
1. Download Microsoft ODBC Driver for SQL Server
2. Install with administrator privileges
3. Verify installation:
   ```bash
   odbcconf -a -s "SQL Server Driver=ODBC Driver 17 for SQL Server"
   ```

#### Linux (Ubuntu/Debian)
```bash
# Add Microsoft repository
curl https://packages.microsoft.com/keys/microsoft.asc | sudo apt-key add -
curl https://packages.microsoft.com/config/ubuntu/20.04/prod.list | sudo tee /etc/apt/sources.list.d/mssql-release.list

# Install driver
sudo apt-get update
sudo apt-get install -y msodbcsql17
```

## Configuration Validation

### 1. Database Connection Test
```python
# test_db_connection.py
import pyodbc
from dotenv import load_dotenv
import os

load_dotenv()

connection_string = (
    f"DRIVER={{ODBC Driver 17 for SQL Server}};"
    f"SERVER={os.getenv('DB_SERVER')};"
    f"DATABASE={os.getenv('DB_DATABASE')};"
    f"UID={os.getenv('DB_USERNAME')};"
    f"PWD={os.getenv('DB_PASSWORD')}"
)

try:
    conn = pyodbc.connect(connection_string)
    print("Database connection successful!")
    conn.close()
except Exception as e:
    print(f"Database connection failed: {e}")
```

### 2. Email Connection Test
```python
# test_email_connection.py
import imaplib
import ssl
from dotenv import load_dotenv
import os

load_dotenv()

def test_imap_connection():
    try:
        server = os.getenv('EMAIL_SERVER')
        port = int(os.getenv('EMAIL_PORT', 993))
        username = os.getenv('EMAIL_USERNAME')
        password = os.getenv('EMAIL_PASSWORD')
        
        # Create SSL context
        context = ssl.create_default_context()
        
        # Connect to server
        imap = imaplib.IMAP4_SSL(server, port, ssl_context=context)
        
        # Login
        imap.login(username, password)
        
        # List folders
        status, folders = imap.list()
        
        print("Email connection successful!")
        print(f"Available folders: {len(folders)}")
        
        imap.logout()
        
    except Exception as e:
        print(f"Email connection failed: {e}")

if __name__ == "__main__":
    test_imap_connection()
```

### 3. Excel Processing Test
```python
# test_excel_processing.py
import pandas as pd
from services.excel_service import ExcelService
import os

def test_excel_processing():
    # Create test Excel file
    test_data = {
        'Name': ['John', 'Jane', 'Bob'],
        'Score': [85, 92, 78],
        'Grade': ['B', 'A', 'C']
    }
    
    df = pd.DataFrame(test_data)
    test_file = 'test_data.xlsx'
    df.to_excel(test_file, index=False)
    
    # Test reading with ExcelService
    excel_service = ExcelService('./logs')
    
    try:
        read_df = excel_service.read_excel(test_file)
        print("Excel processing test successful!")
        print(f"Read {len(read_df)} rows")
        print(read_df.head())
        
        # Clean up
        os.remove(test_file)
        
    except Exception as e:
        print(f"Excel processing test failed: {e}")

if __name__ == "__main__":
    test_excel_processing()
```

## Security Configuration

### 1. Environment Variables Security
- Never commit .env file to version control
- Use strong, unique passwords
- Rotate passwords regularly
- Use app-specific passwords for email services

### 2. Database Security
- Use least privilege principle
- Enable SQL Server authentication encryption
- Regular security updates
- Backup encryption

### 3. File System Security
- Restrict access to download and log folders
- Implement file scanning for uploaded Excel files
- Regular cleanup of old files

## Performance Optimization

### 1. Database Optimization
```sql
-- Additional indexes for better performance
CREATE INDEX IX_Email_Details_ReceivedDate ON Email_Details(ReceivedDate);
CREATE INDEX IX_Email_Details_Subject ON Email_Details(Subject_Name);

-- Partition large tables if needed (for high volume)
-- Consider archiving old data periodically
```

### 2. Application Configuration
```env
# Performance settings
BATCH_SIZE=1000
MAX_RETRIES=3
CONNECTION_TIMEOUT=30
QUERY_TIMEOUT=300
```

### 3. Memory Management
- Monitor memory usage during processing
- Adjust batch size based on available memory
- Implement memory cleanup for large files

## Monitoring and Logging

### 1. Log Configuration
```env
# Logging settings
LOG_LEVEL=INFO
LOG_MAX_SIZE_MB=100
LOG_BACKUP_COUNT=5
LOG_FORMAT=%(asctime)s - %(name)s - %(levelname)s - %(message)s
```

### 2. Monitoring Metrics
- Email processing count
- Database insertion success rate
- Error frequency and types
- Processing time per email
- Disk space usage

## Backup and Recovery

### 1. Database Backup
```sql
-- Create backup script
BACKUP DATABASE EmailAutomation 
TO DISK = 'C:\Backups\EmailAutomation.bak'
WITH FORMAT, INIT;
```

### 2. Configuration Backup
- Backup .env file
- Export stored procedures
- Document custom configurations

### 3. Recovery Procedures
- Database restoration steps
- Configuration recovery
- Data validation post-recovery

## Troubleshooting Common Issues

### 1. Connection Issues
- Check network connectivity
- Verify firewall settings
- Validate credentials
- Test with different connection methods

### 2. Permission Issues
- Verify database permissions
- Check file system permissions
- Validate email account access

### 3. Performance Issues
- Monitor resource usage
- Check database query performance
- Optimize batch processing
- Review log file sizes

This setup guide provides all necessary information for deploying and configuring the Email Reader Automation System in a production environment.
