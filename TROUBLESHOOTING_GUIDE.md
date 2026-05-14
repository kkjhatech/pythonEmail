# Email Reader System - Troubleshooting Guide

## Table of Contents
1. [Common Error Messages](#common-error-messages)
2. [Database Issues](#database-issues)
3. [Email Connection Issues](#email-connection-issues)
4. [Excel Processing Issues](#excel-processing-issues)
5. [Performance Issues](#performance-issues)
6. [Configuration Issues](#configuration-issues)
7. [Debugging Tools](#debugging-tools)
8. [Recovery Procedures](#recovery-procedures)

## Common Error Messages

### Database Connection Errors

#### Error: `Login failed for user`
**Possible Causes:**
- Incorrect username or password
- User doesn't have database access
- SQL Server Authentication not enabled

**Solutions:**
1. Verify credentials in .env file
2. Enable SQL Server Authentication:
   ```sql
   EXEC sp_configure 'show advanced options', 1;
   RECONFIGURE;
   EXEC sp_configure 'SQL Server Authentication', 1;
   RECONFIGURE;
   ```
3. Grant database access:
   ```sql
   USE EmailAutomation;
   CREATE USER [username] FOR LOGIN [username];
   ALTER ROLE db_owner ADD MEMBER [username];
   ```

#### Error: `Cannot open database "EmailAutomation"`
**Possible Causes:**
- Database doesn't exist
- Incorrect database name
- Permissions issue

**Solutions:**
1. Check database exists:
   ```sql
   SELECT name FROM sys.databases WHERE name = 'EmailAutomation';
   ```
2. Create database if missing:
   ```sql
   CREATE DATABASE EmailAutomation;
   ```

#### Error: `ODBC Driver error`
**Possible Causes:**
- ODBC driver not installed
- Incorrect driver version
- Driver architecture mismatch (32-bit vs 64-bit)

**Solutions:**
1. Install correct ODBC driver:
   - Download from Microsoft website
   - Ensure architecture matches Python installation

2. Test driver installation:
   ```bash
   odbcconf -a -s "SQL Server Driver=ODBC Driver 17 for SQL Server"
   ```

### Email Connection Errors

#### Error: `IMAP connection failed`
**Possible Causes:**
- Incorrect server settings
- Firewall blocking port 993
- IMAP not enabled for email account

**Solutions:**
1. Verify IMAP settings:
   - Gmail: `imap.gmail.com:993`
   - Outlook: `outlook.office365.com:993`

2. Enable IMAP in email settings:
   - Gmail: Settings → Forwarding and POP/IMAP
   - Outlook: Settings → Mail → Sync email

3. Check firewall:
   - Ensure port 993 is open
   - Allow Python through firewall

#### Error: `Authentication failed`
**Possible Causes:**
- Incorrect password
- 2FA enabled without app password
- Account locked

**Solutions:**
1. For Gmail with 2FA:
   - Generate App Password
   - Use App Password instead of regular password

2. For Outlook:
   - Use App Password if 2FA enabled
   - Check account security settings

#### Error: `No emails found`
**Possible Causes:**
- No unread emails with attachments
- Emails already processed
- Folder access issue

**Solutions:**
1. Check email folders:
   ```python
   # List available folders
   status, folders = imap.list()
   for folder in folders:
       print(folder)
   ```

2. Send test email with Excel attachment

### Excel Processing Errors

#### Error: `File not found`
**Possible Causes:**
- Download path incorrect
- File permissions issue
- Antivirus blocking

**Solutions:**
1. Verify download folder exists:
   ```bash
   mkdir downloads
   chmod 755 downloads
   ```

2. Check file permissions

#### Error: `Unsupported file format`
**Possible Causes:**
- File extension not supported
- Corrupted Excel file
- Password-protected file

**Solutions:**
1. Check supported formats: `.xlsx`, `.xls`, `.xlsb`, `.csv`
2. Verify file can be opened manually
3. Remove password protection

#### Error: `Data validation failed`
**Possible Causes:**
- Empty Excel file
- Invalid data types
- Missing required columns

**Solutions:**
1. Check Excel file structure
2. Ensure file has data rows
3. Verify column names don't contain special characters

## Database Issues

### Stored Procedure Errors

#### Error: `Could not find stored procedure 'usp_insert_email'`
**Solutions:**
1. Create stored procedures:
   ```sql
   -- Run the stored procedure creation scripts
   -- See SETUP_REQUIREMENTS.md for complete scripts
   ```

2. Verify procedures exist:
   ```sql
   SELECT name FROM sys.procedures WHERE name LIKE 'usp_%';
   ```

### Table Issues

#### Error: `Invalid object name 'Email_Master'`
**Solutions:**
1. Create tables if missing:
   ```sql
   -- Run table creation scripts
   -- See SETUP_REQUIREMENTS.md for complete scripts
   ```

2. Check table exists:
   ```sql
   SELECT name FROM sys.tables WHERE name IN ('Email_Master', 'Email_Details');
   ```

### Data Integrity Issues

#### Error: `Foreign key constraint violation`
**Possible Causes:**
- Email_Details references non-existent Email_Master_A
- Data inconsistency

**Solutions:**
1. Check data integrity:
   ```sql
   SELECT ed.* FROM Email_Details ed
   LEFT JOIN Email_Master em ON ed.EmailID_N = em.Email_Master_A
   WHERE em.Email_Master_A IS NULL;
   ```

2. Clean orphaned records:
   ```sql
   DELETE FROM Email_Details 
   WHERE EmailID_N NOT IN (SELECT Email_Master_A FROM Email_Master);
   ```

## Email Connection Issues

### IMAP vs COM Connection

#### When to use IMAP:
- Remote server access
- Multiple email accounts
- No Outlook installation required

#### When to use COM:
- Local Outlook installation
- Exchange server with MAPI
- Need for Outlook-specific features

### Connection Pooling Issues

#### Error: `Too many connections`
**Solutions:**
1. Reduce check interval:
   ```env
   CHECK_INTERVAL_MINUTES=60
   ```

2. Implement connection reuse in code

### SSL/TLS Issues

#### Error: `SSL handshake failed`
**Solutions:**
1. Update SSL certificate store
2. Use correct SSL settings:
   ```env
   EMAIL_USE_SSL=true
   EMAIL_USE_TLS=false
   ```

## Excel Processing Issues

### Memory Issues

#### Error: `MemoryError` during large file processing
**Solutions:**
1. Process in chunks:
   ```python
   chunksize = 1000
   for chunk in pd.read_excel(file, chunksize=chunksize):
       process_chunk(chunk)
   ```

2. Increase system memory or use 64-bit Python

### Format Issues

#### Error: `Unsupported Excel format`
**Solutions:**
1. Update `openpyxl` and `xlrd`:
   ```bash
   pip install --upgrade openpyxl xlrd
   ```

2. Convert file to supported format

### Encoding Issues

#### Error: `UnicodeDecodeError`
**Solutions:**
1. Specify encoding:
   ```python
   df = pd.read_excel(file, encoding='utf-8')
   ```

2. Use different encoding if needed

## Performance Issues

### Slow Processing

#### Symptoms:
- Long processing times
- High CPU usage
- Memory leaks

#### Solutions:
1. Optimize batch size:
   ```env
   BATCH_SIZE=500
   ```

2. Add database indexes:
   ```sql
   CREATE INDEX IX_Email_Details_ReceivedDate ON Email_Details(ReceivedDate);
   ```

3. Monitor and optimize queries

### Database Performance

#### Symptoms:
- Slow inserts
- Query timeouts
- High disk I/O

#### Solutions:
1. Use parameterized queries
2. Implement connection pooling
3. Regular database maintenance:
   ```sql
   -- Update statistics
   UPDATE STATISTICS Email_Master;
   UPDATE STATISTICS Email_Details;
   
   -- Rebuild indexes
   ALTER INDEX ALL ON Email_Master REBUILD;
   ALTER INDEX ALL ON Email_Details REBUILD;
   ```

## Configuration Issues

### Environment Variables

#### Error: `Key not found in environment`
**Solutions:**
1. Check .env file exists
2. Verify variable names
3. Load environment variables:
   ```python
   from dotenv import load_dotenv
   load_dotenv()
   ```

### Path Issues

#### Error: `FileNotFoundError`
**Solutions:**
1. Use absolute paths
2. Check directory permissions
3. Create missing directories:
   ```bash
   mkdir -p downloads logs
   ```

## Debugging Tools

### Debug Mode

Run with debug logging:
```bash
python main.py --debug --run-once
```

### Database Debugging

Test database connection:
```python
from database import DatabaseManager
from config import Settings

settings = Settings.from_env()
with DatabaseManager(settings) as db:
    result = db.execute_query("SELECT COUNT(*) FROM Email_Master")
    print(f"Email_Master count: {result}")
```

### Email Debugging

Test email connection:
```python
from services import EmailService
from config import Settings

settings = Settings.from_env()
email_service = EmailService(settings)
emails = email_service.get_emails(limit=1)
print(f"Found {len(emails)} emails")
```

### Excel Debugging

Test Excel processing:
```python
from services import ExcelService

excel_service = ExcelService('./logs')
df = excel_service.read_excel('test_file.xlsx')
print(f"DataFrame shape: {df.shape}")
print(f"Columns: {list(df.columns)}")
```

## Recovery Procedures

### Database Recovery

1. **Stop the application**
2. **Backup current data**:
   ```sql
   BACKUP DATABASE EmailAutomation TO DISK = 'backup_before_fix.bak';
   ```

3. **Identify and fix issues**:
   ```sql
   -- Check for orphaned records
   SELECT ed.* FROM Email_Details ed
   LEFT JOIN Email_Master em ON ed.EmailID_N = em.Email_Master_A
   WHERE em.Email_Master_A IS NULL;
   
   -- Fix orphaned records
   DELETE FROM Email_Details 
   WHERE EmailID_N NOT IN (SELECT Email_Master_A FROM Email_Master);
   ```

4. **Restore from backup if needed**:
   ```sql
   RESTORE DATABASE EmailAutomation FROM DISK = 'backup_before_fix.bak';
   ```

### File System Recovery

1. **Check disk space**
2. **Clean up old files**:
   ```bash
   find downloads/ -name "*.xlsx" -mtime +30 -delete
   find logs/ -name "*.log" -mtime +7 -delete
   ```

3. **Verify permissions**:
   ```bash
   chmod -R 755 downloads logs
   ```

### Configuration Recovery

1. **Restore .env file from backup**
2. **Verify all required variables**:
   ```python
   from config import Settings
   try:
       settings = Settings.from_env()
       print("Configuration loaded successfully")
   except Exception as e:
       print(f"Configuration error: {e}")
   ```

## Preventive Measures

### Regular Maintenance

1. **Daily**:
   - Check log files for errors
   - Monitor processing statistics

2. **Weekly**:
   - Archive old log files
   - Check database size
   - Review error patterns

3. **Monthly**:
   - Database maintenance
   - Update dependencies
   - Security review

### Monitoring

1. **Set up alerts for**:
   - Database connection failures
   - Email processing errors
   - Disk space issues

2. **Key metrics to monitor**:
   - Processing success rate
   - Average processing time
   - Error frequency
   - Resource usage

### Backup Strategy

1. **Database backups**:
   - Daily full backups
   - Hourly transaction log backups
   - Test restore procedures

2. **Configuration backups**:
   - Version control for code
   - Backup .env file
   - Document custom settings

## Contact Support

### When to Contact Support
- Critical system failures
- Data corruption issues
- Security incidents
- Performance degradation

### Information to Provide
1. Error messages and logs
2. Steps to reproduce
3. System configuration
4. Recent changes

### Emergency Contacts
- Database Administrator: [Contact Info]
- System Administrator: [Contact Info]
- Application Support: [Contact Info]

This troubleshooting guide should help resolve most common issues with the Email Reader Automation System.
