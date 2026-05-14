# Email Reader System - Step-by-Step Execution Plan

## Phase 1: Initial Setup

### Step 1: Environment Preparation
1. **Install Python 3.8+**
   ```bash
   python --version  # Verify installation
   ```

2. **Create Virtual Environment**
   ```bash
   python -m venv venv
   venv\Scripts\activate  # Windows
   # or
   source venv/bin/activate  # Linux/Mac
   ```

3. **Install Dependencies**
   ```bash
   pip install -r requirements.txt
   ```

### Step 2: Database Setup
1. **Create SQL Server Database**
   ```sql
   CREATE DATABASE EmailAutomation;
   ```

2. **Create Required Tables**
   ```sql
   -- Email_Master Table
   CREATE TABLE Email_Master (
       Email_Master_A INT IDENTITY(1,1) PRIMARY KEY,
       EmailID NVARCHAR(255) NOT NULL,
       CreatedDate DATETIME DEFAULT GETDATE(),
       CreatedBy NVARCHAR(100)
   );

   -- Email_Details Table
   CREATE TABLE Email_Details (
       Email_Details_A INT IDENTITY(1,1) PRIMARY KEY,
       EmailID_N INT NOT NULL,
       Subject_Name NVARCHAR(50),
       SheetName NVARCHAR(50),
       TotalRows INT,
       ReceivedDate DATETIME,
       FOREIGN KEY (EmailID_N) REFERENCES Email_Master(Email_Master_A)
   );
   ```

3. **Create Stored Procedures**
   ```sql
   -- usp_insert_email
   CREATE PROCEDURE [dbo].[usp_insert_email]
       @Email_ID nvarchar(100),
       @CreatedBy nvarchar(100)
   AS
   BEGIN
       SET NOCOUNT ON;
       INSERT INTO Email_Master (EmailID, CreatedDate, CreatedBy) 
       VALUES (@Email_ID, GETDATE(), @CreatedBy)
   END;

   -- usp_insert_email_details
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
   END;
   ```

### Step 3: Configuration Setup
1. **Create .env File**
   ```env
   # Database Configuration
   DB_SERVER=localhost
   DB_DATABASE=EmailAutomation
   DB_USERNAME=your_username
   DB_PASSWORD=your_password

   # Email Configuration
   EMAIL_PROVIDER=outlook
   EMAIL_SERVER=outlook.office365.com
   EMAIL_PORT=993
   EMAIL_USERNAME=your_email@domain.com
   EMAIL_PASSWORD=your_app_password

   # Outlook COM Configuration
   OUTLOOK_CONNECTION_METHOD=com

   # Processing Configuration
   DOWNLOAD_FOLDER=./downloads
   LOG_FOLDER=./logs
   CHECK_INTERVAL_MINUTES=30

   # Email Processing
   EMAIL_PROCESSED_FOLDER=Processed
   ```

2. **Create Required Directories**
   ```bash
   mkdir downloads
   mkdir logs
   ```

## Phase 2: Testing and Validation

### Step 4: Test Database Connection
1. **Run Connection Test**
   ```python
   from database import DatabaseManager
   from config import Settings
   
   settings = Settings.from_env()
   with DatabaseManager(settings) as db:
       if db.test_connection():
           print("Database connection successful!")
       else:
           print("Database connection failed!")
   ```

### Step 5: Test Email Service
1. **Test Email Connection**
   ```bash
   python main.py --test-connection
   ```

2. **Verify Email Access**
   - Check if emails can be retrieved
   - Verify attachment access

### Step 6: Test Excel Processing
1. **Create Test Excel File**
   - Create a sample Excel file with data
   - Place in test folder

2. **Test Excel Reading**
   ```python
   from services import ExcelService
   
   excel_service = ExcelService('./logs')
   df = excel_service.read_excel('test_file.xlsx')
   print(df.head())
   ```

## Phase 3: Production Deployment

### Step 7: Initial Run
1. **Run Once for Testing**
   ```bash
   python main.py --run-once
   ```

2. **Check Results**
   - Verify Email_Master table has sender emails
   - Verify Email_Details table has email metadata
   - Verify prefixed data tables created (PY_1_2_FileName)

### Step 8: Scheduler Setup
1. **Start Automatic Processing**
   ```bash
   python main.py --start-scheduler
   ```

2. **Monitor Logs**
   ```bash
   tail -f logs/email_automation.log
   ```

### Step 9: Production Validation
1. **Send Test Email with Excel Attachment**
2. **Verify Processing**
   - Check if email is processed
   - Verify data in database
   - Check log files for errors

## Phase 4: Ongoing Maintenance

### Step 10: Monitoring
1. **Daily Checks**
   - Review log files for errors
   - Check database table growth
   - Verify email processing

2. **Weekly Checks**
   - Archive old log files
   - Review processing statistics
   - Check disk space

### Step 11: Troubleshooting
1. **Common Issues**
   - Email connection failures
   - Database connection issues
   - Excel file format problems
   - Permission issues

2. **Debug Mode**
   ```bash
   python main.py --debug --run-once
   ```

## Phase 5: Advanced Configuration

### Step 12: Performance Optimization
1. **Database Indexes**
   ```sql
   CREATE INDEX IX_Email_Master_EmailID ON Email_Master(EmailID);
   CREATE INDEX IX_Email_Details_EmailID_N ON Email_Details(EmailID_N);
   ```

2. **Batch Size Optimization**
   - Adjust batch_size in insert_dataframe method
   - Monitor memory usage

### Step 13: Security Hardening
1. **Credential Management**
   - Use environment variables
   - Rotate passwords regularly
   - Use app-specific passwords

2. **Access Control**
   - Limit database user permissions
   - Secure file system permissions

## Execution Commands Summary

### Development Commands
```bash
# Setup
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt

# Testing
python main.py --test-connection
python main.py --run-once
python main.py --debug --run-once
```

### Production Commands
```bash
# Start scheduler
python main.py --start-scheduler

# Run once
python main.py --run-once

# Stop gracefully
Ctrl+C
```

## Verification Checklist

### Pre-Deployment
- [ ] Python 3.8+ installed
- [ ] Virtual environment created
- [ ] Dependencies installed
- [ ] Database created
- [ ] Tables created
- [ ] Stored procedures created
- [ ] .env file configured
- [ ] Directories created

### Post-Deployment
- [ ] Database connection working
- [ ] Email service connected
- [ ] Excel files processed
- [ ] Data inserted correctly
- [ ] Scheduler running
- [ ] Logs being generated
- [ ] No critical errors

### Ongoing
- [ ] Daily log review
- [ ] Weekly statistics check
- [ ] Monthly performance review
- [ ] Quarterly security audit

## Rollback Plan

### If Issues Occur
1. **Stop the Application**
   ```bash
   Ctrl+C  # Stop scheduler
   ```

2. **Database Rollback**
   ```sql
   -- Drop prefixed tables
   DROP TABLE IF EXISTS PY_1_2_TestFile;
   
   -- Clear new entries
   DELETE FROM Email_Details WHERE ReceivedDate > 'rollback_date';
   DELETE FROM Email_Master WHERE CreatedDate > 'rollback_date';
   ```

3. **Configuration Rollback**
   - Restore previous .env file
   - Check previous version in source control

## Support Contacts

### Technical Support
- Database Administrator: For SQL Server issues
- Email Administrator: For email server problems
- System Administrator: For server and deployment issues

### Emergency Contacts
- Primary: [Contact Information]
- Secondary: [Contact Information]
- On-call: [Contact Information]
