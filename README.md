# Email Automation System

A production-ready Python automation system that connects to email inboxes, processes Excel attachments, and inserts data into SQL Server databases.

## Features

- **Email Integration**: Connects to Gmail via IMAP (Microsoft 365 support available)
- **Smart Filtering**: Filters by sender, subject keywords, and date range
- **Excel Processing**: Reads `.xlsx`, `.xls`, and `.csv` files using Pandas/OpenPyXL
- **Data Validation**: Validates columns and data types before database insertion
- **Duplicate Prevention**: Prevents reprocessing the same emails/data
- **SQL Server Integration**: Dynamic table creation and data insertion via pyodbc
- **Comprehensive Logging**: Detailed logs for all operations
- **Scheduler**: Automatic execution at configurable intervals
- **Secure**: Environment-based credential storage

## Project Structure

```
Email_Reader/
├── config/
│   ├── __init__.py
│   └── settings.py           # Configuration management
├── services/
│   ├── __init__.py
│   ├── email_service.py      # Email IMAP operations
│   ├── excel_service.py      # Excel file processing
│   └── scheduler_service.py  # Task scheduling
├── database/
│   ├── __init__.py
│   └── db_manager.py         # SQL Server operations
├── utils/
│   ├── __init__.py
│   ├── logger.py             # Logging configuration
│   └── validators.py         # Data validation
├── downloads/                # Downloaded attachments (ignored by git)
├── logs/                     # Application logs (ignored by git)
├── .env                      # Environment variables (ignored by git)
├── .env.example              # Example environment file
├── .gitignore
├── main.py                   # Application entry point
├── requirements.txt          # Python dependencies
├── setup.py                  # Package setup
├── sql_setup.sql             # Database setup script
└── README.md                 # This file
```

## Prerequisites

1. **Python**: Version 3.9 or higher
2. **SQL Server**: Local or remote instance
3. **Gmail Account**: With App Password enabled
4. **ODBC Driver**: SQL Server ODBC Driver 17 or higher

## Installation

### 1. Clone/Extract the Project

```bash
cd Email_Reader
```

### 2. Create Virtual Environment

```bash
python -m venv venv

# Windows
venv\Scripts\activate

# macOS/Linux
source venv/bin/activate
```

### 3. Install Dependencies

```bash
pip install -r requirements.txt
```

### 4. Install SQL Server ODBC Driver (if not installed)

**Windows**: Download from [Microsoft](https://docs.microsoft.com/en-us/sql/connect/odbc/download-odbc-driver-for-sql-server)

**Ubuntu/Debian**:
```bash
curl https://packages.microsoft.com/keys/microsoft.asc | apt-key add -
curl https://packages.microsoft.com/config/ubuntu/$(lsb_release -rs)/prod.list > /etc/apt/sources.list.d/mssql-release.list
apt-get update
ACCEPT_EULA=Y apt-get install -y msodbcsql17
```

### 5. Configure Environment Variables

Copy the example file and fill in your credentials:

```bash
cp .env.example .env
```

Edit `.env` with your settings:

```env
# Email Configuration (Gmail)
EMAIL_HOST=imap.gmail.com
EMAIL_PORT=993
EMAIL_USERNAME=your-email@gmail.com
EMAIL_PASSWORD=your-gmail-app-password
EMAIL_USE_SSL=true

# Mailbox Configuration
EMAIL_INBOX_FOLDER=INBOX
EMAIL_PROCESSED_FOLDER=Processed      # Leave empty to only mark as read

# Database Configuration
DB_SERVER=localhost
DB_NAME=EmailAutomationDB
DB_USERNAME=sa
DB_PASSWORD=your-password
DB_DRIVER=ODBC Driver 17 for SQL Server

# Application Configuration
DOWNLOAD_FOLDER=downloads
LOG_FOLDER=logs
CHECK_INTERVAL_MINUTES=5
DATE_FILTER_DAYS=7

# Filter Configuration
ALLOWED_SENDERS=trusted@company.com,reports@partner.com
SUBJECT_KEYWORDS=Report,Data,Excel,Monthly
FILE_EXTENSIONS=.xlsx,.xls,.csv
```

**Note**: For Gmail, use an [App Password](https://support.google.com/accounts/answer/185833), not your regular password.

### 6. Setup Database

Run the SQL setup script:

```bash
sqlcmd -S localhost -U sa -P your-password -i sql_setup.sql
```

Or execute `sql_setup.sql` in SQL Server Management Studio (SSMS).

## Usage

### Run Once (Manual Mode)

Process emails immediately and exit:

```bash
python main.py --run-once
```

### Run with Scheduler (Automatic Mode)

Start the scheduler for continuous processing:

```bash
python main.py
```

The scheduler will check for new emails every `CHECK_INTERVAL_MINUTES` (default: 5 minutes).

### Test Connections

Test database connection:

```bash
python main.py --test-db
```

Test email connection:

```bash
python main.py --test-email
```

## Configuration Details

### Email Filters

Configure in `.env`:

- **ALLOWED_SENDERS**: Only process emails from these addresses (comma-separated). Leave empty to allow all.
- **SUBJECT_KEYWORDS**: Only process emails with these keywords in subject (comma-separated). Leave empty to allow all.
- **DATE_FILTER_DAYS**: Only process emails from the last N days.
- **FILE_EXTENSIONS**: Acceptable attachment types (default: `.xlsx,.xls,.csv`)

### Database Behavior

- **Dynamic Tables**: Tables are created automatically based on Excel file structure
- **Duplicate Prevention**: Uses `email_id` column to track processed emails
- **Tracking Columns**: Each table includes `email_id` and `processed_date` columns
- **Table Naming**: Based on Excel filename (sanitized for SQL)

### Post-Processing

After processing, emails can be:
1. **Moved to folder**: Set `EMAIL_PROCESSED_FOLDER` (e.g., `Processed`)
2. **Marked as read**: If folder move fails or is not configured

## Logging

Logs are stored in the `logs/` folder with format:
```
YYYY-MM-DD HH:MM:SS | LoggerName | LEVEL | Message
```

Log files are rotated daily with `.log` extension.

### Sample Log Output

```
2024-01-15 09:30:15 | EmailAutomation | INFO | Starting email automation cycle
2024-01-15 09:30:16 | EmailService | INFO | Found 3 unread emails
2024-01-15 09:30:17 | EmailService | INFO | Filtered to 2 matching emails
2024-01-15 09:30:18 | ExcelService | INFO | Successfully read data.xlsx: 150 rows, 5 columns
2024-01-15 09:30:20 | DatabaseManager | INFO | Successfully inserted 150 rows into data
2024-01-15 09:30:21 | EmailService | INFO | Marked email 12345 as read
2024-01-15 09:30:22 | EmailAutomation | INFO | Processing Summary
2024-01-15 09:30:22 | EmailAutomation | INFO | Emails processed: 2
2024-01-15 09:30:22 | EmailAutomation | INFO | Files downloaded: 2
2024-01-15 09:30:22 | EmailAutomation | INFO | Rows inserted: 300
```

## Database Schema

The system creates three types of tables:

1. **Dynamic Data Tables**: Created from Excel files (e.g., `SalesReport_2024`)
2. **processing_log**: Tracks all email processing attempts
3. **import_details**: Detailed import metadata

### Processing Log Table

| Column | Type | Description |
|--------|------|-------------|
| id | INT (PK) | Auto-increment ID |
| email_id | NVARCHAR(255) | Unique email identifier |
| sender_email | NVARCHAR(255) | Sender email address |
| subject | NVARCHAR(500) | Email subject |
| file_name | NVARCHAR(500) | Processed file name |
| table_name | NVARCHAR(100) | Target database table |
| rows_processed | INT | Number of rows inserted |
| status | NVARCHAR(50) | SUCCESS, FAILED, SKIPPED |
| error_message | NVARCHAR(MAX) | Error details if failed |
| processed_date | DATETIME | Processing timestamp |

## Troubleshooting

### Common Issues

#### "Failed to connect to email server"
- Verify IMAP is enabled in Gmail settings
- Use App Password, not regular Gmail password
- Check firewall/antivirus isn't blocking port 993

#### "Failed to connect to database"
- Verify SQL Server is running
- Check ODBC Driver 17 is installed
- Verify connection string in `.env`
- Enable SQL Server authentication (mixed mode)

#### "No module named 'pyodbc'"
```bash
pip install pyodbc
```

On Linux, you may need build dependencies:
```bash
sudo apt-get install unixodbc-dev
```

#### "Data validation failed"
- Check Excel file isn't corrupted
- Ensure columns have headers
- Verify data types are consistent

### Debug Mode

Enable debug logging by modifying `utils/logger.py`:

```python
logger.setLevel(logging.DEBUG)
console_handler.setLevel(logging.DEBUG)
```

## Security Considerations

1. **Never commit `.env` file** - It contains sensitive credentials
2. **Use App Passwords** for Gmail instead of main password
3. **Restrict SQL permissions** - Use dedicated DB user with minimal privileges
4. **Secure download folder** - Ensure proper file permissions
5. **Enable SSL/TLS** for email connections (enabled by default)

## Extending the System

### Adding New File Types

Edit `settings.py` to add new extensions:

```python
FILE_EXTENSIONS=.xlsx,.xls,.csv,.json
```

Then modify `excel_service.py` to handle the new format.

### Custom Data Transformation

Extend `ExcelService.validate_and_prepare()` to add custom transformations:

```python
def custom_transform(self, df: pd.DataFrame) -> pd.DataFrame:
    # Add your transformation logic
    df['processed_at'] = datetime.now()
    return df
```

### Microsoft 365 Support

To use Microsoft Graph API instead of IMAP:

1. Register app in Azure AD
2. Install `O365` package: `pip install O365`
3. Replace `EmailService` with Graph API implementation

## License

MIT License - See LICENSE file for details.

## Support

For issues or questions:
1. Check the logs in `logs/` folder
2. Verify configuration in `.env`
3. Test connections with `--test-db` and `--test-email`

---

**Production Note**: Before deploying to production:
- Set up proper monitoring and alerting
- Configure log rotation
- Implement backup strategy for downloaded files
- Review and tighten database permissions
- Test thoroughly in staging environment
