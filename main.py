"""
Email Automation System
Main entry point for email processing automation.
"""

import sys
import signal
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Any, Optional, Tuple

from config import Settings
from services import EmailService, ExcelService, SchedulerService, OutlookCOMAutoService
from database import DatabaseManager
from utils import get_logger, ExcelValidator


class EmailAutomation:
    """Main application class for email automation."""
    
    def __init__(self):
        self.settings = Settings.from_env()
        self.logger = get_logger('EmailAutomation', self.settings.log_folder)
        self.scheduler = SchedulerService(self.settings.log_folder)
        self.excel_service = ExcelService(self.settings.log_folder)
        
        # Statistics
        self.stats = {
            'emails_processed': 0,
            'files_downloaded': 0,
            'rows_inserted': 0,
            'errors': 0
        }
    
    def _get_email_service(self):
        """Get appropriate email service based on provider and connection method."""
        provider = self.settings.email_provider
        
        if provider == 'gmail':
            self.logger.info("Using Gmail IMAP service")
            return EmailService(self.settings)
        
        elif provider in ['outlook', 'office365']:
            if self.settings.outlook_connection_method == 'com':
                self.logger.info("Using Outlook COM automation (requires Outlook desktop app)")
                return OutlookCOMAutoService(self.settings)
            else:
                self.logger.info("Using Outlook IMAP service")
                return EmailService(self.settings)
        
        else:
            self.logger.info(f"Using default IMAP service for {provider}")
            return EmailService(self.settings)
    
    def run(self):
        """Run the automation once."""
        self.logger.info("=" * 60)
        self.logger.info("Starting email automation cycle")
        self.logger.info("=" * 60)
        
        try:
            with self._get_email_service() as email_service:
                # Fetch unread emails
                emails = email_service.fetch_unread_emails(
                    since_days=self.settings.date_filter_days
                )
                
                if not emails:
                    self.logger.info("No matching emails found")
                    return
                
                self.logger.info(f"Processing {len(emails)} emails")
                
                # Process each email
                for email_data in emails:
                    try:
                        self._process_email(email_service, email_data)
                    except Exception as e:
                        self.logger.error(f"Error processing email: {str(e)}")
                        self.stats['errors'] += 1
                
                # Print summary
                self._print_summary()
                
        except Exception as e:
            self.logger.error(f"Automation cycle failed: {str(e)}")
            self.stats['errors'] += 1
    
    def _process_email(
        self,
        email_service: EmailService,
        email_data: Dict[str, Any]
    ):
        """Process a single email."""
        email_id = email_data['id']
        sender = email_data.get('sender_email', 'unknown')
        subject = email_data.get('subject', 'no subject')
        
        self.logger.info(f"Processing email from {sender}: {subject}")
        
        # Check for Excel attachments
        if not email_service.has_excel_attachments(email_data):
            self.logger.info("No Excel attachments found, skipping")
            self._post_process_email(email_service, email_data, processed=False)
            return
        
        # Process attachments
        files_processed = 0
        for attachment in email_data.get('attachments', []):
            filename = attachment['filename'].lower()
            
            # Check if it's an allowed file type
            if not any(filename.endswith(ext) for ext in self.settings.file_extensions):
                continue
            
            # Download file
            file_path = email_service.download_attachment(
                attachment,
                self.settings.download_folder
            )
            
            if file_path:
                self.stats['files_downloaded'] += 1
                self._process_excel_file(file_path, email_data, sender)
                files_processed += 1
        
        # Post-process email
        if files_processed > 0:
            self.stats['emails_processed'] += 1
            self._post_process_email(email_service, email_data, processed=True)
        else:
            self._post_process_email(email_service, email_data, processed=False)
    
    def _process_excel_file(self, file_path: str, email_data: Dict[str, Any], sender: str):
        """Process a downloaded Excel file."""
        self.logger.info(f"Processing Excel file: {file_path}")
        
        try:
            # Read Excel file
            df = self.excel_service.read_excel(file_path)
            
            if df is None:
                self.logger.error(f"Failed to read file: {file_path}")
                return
            
            # Get table name from filename
            table_name = Path(file_path).stem
            
            # Validate and prepare data
            is_valid, prepared_df, message = self.excel_service.validate_and_prepare(
                df,
                table_name
            )
            
            if not is_valid:
                self.logger.error(f"Data validation failed: {message}")
                return
            
            # DEBUG: Check prepared data
            self.logger.info(f"Prepared data dtypes:\n{prepared_df.dtypes}")
            self.logger.info(f"Prepared data first row:\n{prepared_df.iloc[0].to_dict()}")
            
            # Insert into database with new structure
            with DatabaseManager(self.settings) as db:
                # Step 1: Insert into Email_Master table using sender email
                sender_email = email_data.get('sender_email', '')
                if not sender_email:
                    # Extract email from sender string if sender_email is not available
                    sender_str = email_data.get('sender', '')
                    if '<' in sender_str and '>' in sender_str:
                        sender_email = sender_str.split('<')[1].split('>')[0].lower()
                    else:
                        sender_email = sender_str.lower()
                
                self.logger.info(f"DEBUG: email_data keys: {list(email_data.keys())}")
                self.logger.info(f"DEBUG: sender_email being inserted: '{sender_email}'")
                success, email_master_a, msg = db.insert_email_master(sender_email, "System")
                
                if not success:
                    self.logger.error(f"Failed to insert into Email_Master: {msg}")
                    self.stats['errors'] += 1
                    return
                
                # Step 2: Insert into Email_Details table
                subject = email_data.get('subject', '')
                sheet_name = 'Sheet1'  # Default, could be enhanced to detect actual sheet name
                total_rows = len(prepared_df)
                received_date = email_data.get('date', datetime.now())
                
                success, email_details_a, msg = db.insert_email_details(
                    email_master_a, subject, sheet_name, total_rows, received_date
                )
                
                if not success:
                    self.logger.error(f"Failed to insert into Email_Details: {msg}")
                    self.stats['errors'] += 1
                    return
                
                # Step 3: Create data table with prefixed name
                prefixed_table_name = f"PY_{email_master_a}_{email_details_a}_{table_name}"
                
                if not db.table_exists(prefixed_table_name):
                    self.logger.info(f"Table {prefixed_table_name} doesn't exist, creating...")
                    create_sql = self.excel_service.generate_create_table_sql(
                        table_name,
                        prepared_df,
                        email_master_a,
                        email_details_a
                    )
                    # DEBUG: Log CREATE TABLE SQL
                    self.logger.info(f"CREATE TABLE SQL:\n{create_sql}")
                    db.execute_query(create_sql)
                    self.logger.info(f"Table {prefixed_table_name} created successfully")
                
                # Step 4: Insert data into prefixed table
                success, rows, message = db.insert_dataframe(
                    prepared_df,
                    prefixed_table_name,
                    sender,
                    email_details_a
                )
                
                if success:
                    self.stats['rows_inserted'] += rows
                    self.logger.info(f"Data inserted into {prefixed_table_name}: {message}")
                else:
                    self.logger.error(f"Data insertion failed: {message}")
                    self.stats['errors'] += 1
            
        except Exception as e:
            self.logger.error(f"Error processing Excel file {file_path}: {str(e)}")
            self.stats['errors'] += 1
    
    def _generate_table_name(self, file_path: str) -> str:
        """Generate a table name from file path."""
        path = Path(file_path)
        name = path.stem
        
        # Remove timestamp suffix (format: _YYYYMMDD_HHMMSS)
        import re
        name = re.sub(r'_\d{8}_\d{6}$', '', name)
        
        # Sanitize for SQL
        name = name.replace(' ', '_').replace('-', '_')
        name = re.sub(r'[^a-zA-Z0-9_]', '', name)
        
        # Remove leading digits
        if name and name[0].isdigit():
            name = 'tbl_' + name
        
        # Ensure valid length
        if len(name) > 100:
            name = name[:100]
        
        return name or 'imported_data'
    
    def _post_process_email(
        self,
        email_service: EmailService,
        email_data: Dict[str, Any],
        processed: bool = True
    ):
        """Handle post-processing of email (mark read or move folder)."""
        email_id = email_data['id']
        
        try:
            # Try to move to processed folder first
            if self.settings.email_processed_folder:
                if email_service.move_to_folder(
                    email_data,
                    self.settings.email_processed_folder
                ):
                    return
            
            # Fall back to marking as read
            email_service.mark_as_read(email_data)
            
        except Exception as e:
            self.logger.warning(f"Post-processing failed for email {email_id}: {str(e)}")
    
    def _print_summary(self):
        """Print processing summary."""
        self.logger.info("=" * 60)
        self.logger.info("Processing Summary")
        self.logger.info("=" * 60)
        self.logger.info(f"Emails processed: {self.stats['emails_processed']}")
        self.logger.info(f"Files downloaded: {self.stats['files_downloaded']}")
        self.logger.info(f"Rows inserted: {self.stats['rows_inserted']}")
        self.logger.info(f"Errors: {self.stats['errors']}")
        self.logger.info("=" * 60)
    
    def start_scheduler(self):
        """Start the scheduler for automatic execution."""
        self.logger.info(f"Starting scheduler (interval: {self.settings.check_interval_minutes} minutes)")
        
        # Set up signal handlers for graceful shutdown
        signal.signal(signal.SIGINT, self._signal_handler)
        signal.signal(signal.SIGTERM, self._signal_handler)
        
        # Start scheduler
        self.scheduler.start(
            task=self.run,
            interval_minutes=self.settings.check_interval_minutes,
            run_immediately=True
        )
        
        # Keep main thread alive
        try:
            while self.scheduler.is_running():
                import time
                time.sleep(1)
        except KeyboardInterrupt:
            self.stop_scheduler()
    
    def stop_scheduler(self):
        """Stop the scheduler."""
        self.logger.info("Stopping scheduler...")
        self.scheduler.stop()
        self.logger.info("Scheduler stopped")
    
    def _signal_handler(self, signum, frame):
        """Handle shutdown signals."""
        self.logger.info(f"Received signal {signum}, shutting down...")
        self.stop_scheduler()
        sys.exit(0)


def main():
    """Main entry point."""
    import argparse
    
    parser = argparse.ArgumentParser(description='Email Automation System')
    parser.add_argument(
        '--run-once',
        action='store_true',
        help='Run once and exit (no scheduler)'
    )
    parser.add_argument(
        '--test-db',
        action='store_true',
        help='Test database connection and exit'
    )
    parser.add_argument(
        '--test-email',
        action='store_true',
        help='Test email connection and exit'
    )
    
    args = parser.parse_args()
    
    settings = Settings.from_env()
    
    # Test database connection
    if args.test_db:
        logger = get_logger('Main', settings.log_folder)
        logger.info("Testing database connection...")
        with DatabaseManager(settings) as db:
            if db.test_connection():
                logger.info("Database connection successful!")
                return 0
            else:
                logger.error("Database connection failed!")
                return 1
    
    # Test email connection
    if args.test_email:
        logger = get_logger('Main', settings.log_folder)
        logger.info("Testing email connection...")
        with EmailService(settings) as email:
            if email.connect():
                logger.info("Email connection successful!")
                return 0
            else:
                logger.error("Email connection failed!")
                return 1
    
    # Run automation
    automation = EmailAutomation()
    
    if args.run_once:
        automation.run()
    else:
        automation.start_scheduler()
    
    return 0


if __name__ == '__main__':
    sys.exit(main())
