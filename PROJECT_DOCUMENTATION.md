# Email Reader Automation System

## Project Overview

The Email Reader Automation System is a Python application that automatically processes emails with Excel attachments, extracts data from the Excel files, and stores the data in a SQL Server database with proper email tracking and relationships.

## Architecture

### Core Components

1. **EmailAutomation** (`main.py`)
   - Main orchestrator class
   - Manages email processing workflow
   - Handles scheduler for automatic execution

2. **Email Services**
   - `EmailService` - IMAP-based email processing (Gmail, Outlook IMAP)
   - `OutlookCOMAutoService` - Outlook COM automation (desktop app)

3. **Excel Service** (`services/excel_service.py`)
   - Reads Excel files (.csv, .xlsx, .xls, .xlsb)
   - Validates and prepares data for database insertion
   - Generates CREATE TABLE SQL statements

4. **Database Manager** (`database/db_manager.py`)
   - Handles SQL Server connections
   - Manages data insertion with tracking
   - Implements new Email_Master and Email_Details tables

5. **Utilities** (`utils/`)
   - `validators.py` - Data validation and column sanitization
   - Logging configuration

## Database Schema

### New Table Structure

#### 1. Email_Master Table
```sql
CREATE TABLE Email_Master (
    Email_Master_A INT IDENTITY(1,1) PRIMARY KEY,
    EmailID NVARCHAR(255) NOT NULL,        -- Actual sender email
    CreatedDate DATETIME DEFAULT GETDATE(),
    CreatedBy NVARCHAR(100)
);
```

#### 2. Email_Details Table
```sql
CREATE TABLE Email_Details (
    Email_Details_A INT IDENTITY(1,1) PRIMARY KEY,
    EmailID_N INT NOT NULL,                -- Foreign key to Email_Master_A
    Subject_Name NVARCHAR(50),
    SheetName NVARCHAR(50),
    TotalRows INT,
    ReceivedDate DATETIME,
    FOREIGN KEY (EmailID_N) REFERENCES Email_Master(Email_Master_A)
);
```

#### 3. Prefixed Data Tables
```sql
-- Format: PY_{Email_Master_A}_{Email_Details_A}_{FileName}
CREATE TABLE PY_1_2_SubjectMarks (
    id INT IDENTITY(1,1) PRIMARY KEY,
    [Email_Details_A] INT,                -- For join purposes
    processed_date DATETIME DEFAULT GETDATE(),
    [sno] NVARCHAR(500),
    [Student_name] NVARCHAR(500),
    [phy] NVARCHAR(500),
    [che] NVARCHAR(500),
    [math] NVARCHAR(500),
    [bio] NVARCHAR(500),
    [english] NVARCHAR(500),
    [Total] NVARCHAR(500)
);
```

## Data Flow

1. **Email Retrieval**
   - Connects to email server (IMAP or Outlook COM)
   - Fetches unread emails with Excel attachments

2. **Email Processing**
   - Downloads Excel attachments to local folder
   - Extracts email metadata (sender, subject, date)

3. **Database Insertion**
   - Step 1: Insert sender email into Email_Master (checks duplicates)
   - Step 2: Insert email details into Email_Details (linked to Email_Master_A)
   - Step 3: Create prefixed data table with format `PY_1_2_FileName`
   - Step 4: Insert Excel data into prefixed table with Email_Details_A column

4. **Post-Processing**
   - Marks email as read or moves to processed folder
   - Updates processing statistics

## Key Features

- **Duplicate Prevention**: Checks Email_Master for existing sender emails
- **Relational Integrity**: Email_Details linked to Email_Master via foreign key
- **Prefixed Tables**: Data tables named `PY_{Email_Master_A}_{Email_Details_A}_{FileName}`
- **Join Capability**: Email_Details_A column in data tables for easy joins
- **Multiple Email Providers**: Supports Gmail, Outlook IMAP, and Outlook COM
- **Scheduled Execution**: Automatic processing at configurable intervals
- **Error Handling**: Comprehensive error logging and recovery

## Configuration

### Environment Variables (.env)
```env
# Database Configuration
DB_SERVER=your_server_name
DB_DATABASE=your_database_name
DB_USERNAME=your_username
DB_PASSWORD=your_password

# Email Configuration
EMAIL_PROVIDER=outlook          # gmail, outlook, office365
EMAIL_SERVER=outlook.office365.com
EMAIL_PORT=993
EMAIL_USERNAME=your_email@domain.com
EMAIL_PASSWORD=your_app_password

# Outlook COM Configuration
OUTLOOK_CONNECTION_METHOD=com   # com or imap

# Processing Configuration
DOWNLOAD_FOLDER=./downloads
LOG_FOLDER=./logs
CHECK_INTERVAL_MINUTES=30

# Email Processing
EMAIL_PROCESSED_FOLDER=Processed
```

## Stored Procedures

### 1. usp_insert_email
```sql
CREATE PROCEDURE [dbo].[usp_insert_email]
    @Email_ID nvarchar(100),
    @CreatedBy nvarchar(100)
AS
BEGIN
    SET NOCOUNT ON;
    INSERT INTO Email_Master (EmailID, CreatedDate, CreatedBy) 
    VALUES (@Email_ID, GETDATE(), @CreatedBy)
END
```

### 2. usp_insert_email_details
```sql
CREATE PROCEDURE [dbo].[usp_insert_email_details]
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
END
```

## Dependencies

### Python Packages
```
pyodbc>=4.0.32
pandas>=1.5.0
openpyxl>=3.0.0
xlrd>=2.0.0
python-dotenv>=0.19.0
```

### System Requirements
- Python 3.8+
- SQL Server 2016+
- Microsoft Outlook (for COM automation)
- ODBC Driver for SQL Server

## File Structure

```
Email_Reader/
├── main.py                     # Main application entry point
├── config.py                   # Configuration management
├── .env                        # Environment variables
├── requirements.txt            # Python dependencies
├── database/
│   └── db_manager.py          # Database operations
├── services/
│   ├── email_service.py       # IMAP email service
│   ├── outlook_com_service.py # Outlook COM service
│   ├── excel_service.py       # Excel processing
│   └── scheduler_service.py   # Task scheduling
├── utils/
│   └── validators.py          # Data validation
└── logs/                      # Log files directory
```

## Security Considerations

1. **Email Credentials**: Use app-specific passwords for Gmail
2. **Database Security**: Use SQL Server authentication with encrypted passwords
3. **File Access**: Ensure proper permissions for download and log folders
4. **Data Privacy**: Email content is processed locally; no external API calls

## Performance Optimization

1. **Batch Processing**: Inserts data in batches of 1000 rows
2. **Connection Pooling**: Reuses database connections
3. **Indexing**: Consider indexes on Email_Master.EmailID and Email_Details.EmailID_N
4. **Memory Management**: Processes one email at a time to avoid memory issues

## Monitoring and Maintenance

1. **Log Files**: Check logs/ folder for processing logs
2. **Database Monitoring**: Monitor table growth and performance
3. **Error Alerts**: Implement email alerts for critical errors
4. **Regular Cleanup**: Archive old processed emails and data
