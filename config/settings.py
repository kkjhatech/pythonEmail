import os
from dataclasses import dataclass
from pathlib import Path
from typing import List
from dotenv import load_dotenv


@dataclass
class Settings:
    """Application configuration settings loaded from environment variables."""
    
    # Email Configuration
    email_host: str
    email_port: int
    email_username: str
    email_password: str
    email_use_ssl: bool
    email_provider: str  # 'gmail' or 'outlook'
    
    # Mailbox Configuration
    email_inbox_folder: str
    email_processed_folder: str
    
    # Database Configuration
    db_server: str
    db_name: str
    db_username: str
    db_password: str
    db_driver: str
    
    # Application Configuration
    download_folder: str
    log_folder: str
    check_interval_minutes: int
    date_filter_days: int
    
    # Filter Configuration
    allowed_senders: List[str]
    subject_keywords: List[str]
    file_extensions: List[str]
    
    @classmethod
    def from_env(cls) -> 'Settings':
        """Load settings from environment variables."""
        load_dotenv()
        
        # Base paths
        base_dir = Path(__file__).parent.parent
        
        return cls(
            # Email Provider
            email_provider=os.getenv('EMAIL_PROVIDER', 'gmail').lower(),
            
            # Email settings based on provider
            email_host=os.getenv('EMAIL_HOST', cls._get_default_host(os.getenv('EMAIL_PROVIDER', 'gmail'))),
            email_port=int(os.getenv('EMAIL_PORT', '993')),
            email_username=os.getenv('EMAIL_USERNAME', ''),
            email_password=os.getenv('EMAIL_PASSWORD', ''),
            email_use_ssl=os.getenv('EMAIL_USE_SSL', 'true').lower() == 'true',
            
            # Mailbox
            email_inbox_folder=os.getenv('EMAIL_INBOX_FOLDER', 'INBOX'),
            email_processed_folder=os.getenv('EMAIL_PROCESSED_FOLDER', 'Processed'),
            
            # Database
            db_server=os.getenv('DB_SERVER', 'localhost'),
            db_name=os.getenv('DB_NAME', 'EmailAutomationDB'),
            db_username=os.getenv('DB_USERNAME', ''),
            db_password=os.getenv('DB_PASSWORD', ''),
            db_driver=os.getenv('DB_DRIVER', 'ODBC Driver 17 for SQL Server'),
            
            # Application
            download_folder=os.getenv('DOWNLOAD_FOLDER', str(base_dir / 'downloads')),
            log_folder=os.getenv('LOG_FOLDER', str(base_dir / 'logs')),
            check_interval_minutes=int(os.getenv('CHECK_INTERVAL_MINUTES', '5')),
            date_filter_days=int(os.getenv('DATE_FILTER_DAYS', '0')),  # 0 = no date filter
            
            # Filters
            allowed_senders=cls._parse_list(os.getenv('ALLOWED_SENDERS', '')),
            subject_keywords=cls._parse_list(os.getenv('SUBJECT_KEYWORDS', '')),
            file_extensions=cls._parse_list(os.getenv('FILE_EXTENSIONS', '.xlsx,.xls,.csv')),
        )
    
    @staticmethod
    def _parse_list(value: str) -> List[str]:
        """Parse comma-separated string into list."""
        if not value:
            return []
        return [item.strip() for item in value.split(',') if item.strip()]
    
    @staticmethod
    def _get_default_host(provider: str) -> str:
        """Get default IMAP host based on provider."""
        hosts = {
            'gmail': 'imap.gmail.com',
            'outlook': 'outlook.office365.com',
            'hotmail': 'outlook.office365.com',
            'office365': 'outlook.office365.com',
            'yahoo': 'imap.mail.yahoo.com',
        }
        return hosts.get(provider.lower(), 'imap.gmail.com')
    
    def get_db_connection_string(self) -> str:
        """Generate SQL Server connection string."""
        return (
            f"DRIVER={{{self.db_driver}}};"
            f"SERVER={self.db_server};"
            f"DATABASE={self.db_name};"
            f"UID={self.db_username};"
            f"PWD={self.db_password};"
            "TrustServerCertificate=yes;"
        )
