import imaplib
import email
from email.message import Message
from email.header import decode_header
from email.utils import parsedate_to_datetime
from pathlib import Path
from datetime import datetime, timedelta, timezone
from typing import List, Optional, Tuple, Dict, Any
import re

from config.settings import Settings
from utils.logger import get_logger


class EmailService:
    """Service for handling email operations via IMAP."""
    
    def __init__(self, settings: Settings):
        self.settings = settings
        self.logger = get_logger('EmailService', settings.log_folder)
        self.connection: Optional[imaplib.IMAP4_SSL] = None
    
    def connect(self) -> bool:
        """Establish connection to email server."""
        try:
            self.logger.info(f"Connecting to {self.settings.email_provider} ({self.settings.email_host}:{self.settings.email_port})")
            
            if self.settings.email_use_ssl:
                self.connection = imaplib.IMAP4_SSL(
                    self.settings.email_host,
                    self.settings.email_port
                )
            else:
                self.connection = imaplib.IMAP4(
                    self.settings.email_host,
                    self.settings.email_port
                )
            
            # Login
            self.connection.login(
                self.settings.email_username,
                self.settings.email_password
            )
            
            self.logger.info("Email connection established successfully")
            return True
            
        except Exception as e:
            self.logger.error(f"Failed to connect to email server: {str(e)}")
            return False
    
    def disconnect(self):
        """Close email connection."""
        if self.connection:
            try:
                self.connection.close()
                self.connection.logout()
                self.logger.info("Email connection closed")
            except Exception as e:
                self.logger.warning(f"Error closing connection: {str(e)}")
            finally:
                self.connection = None
    
    def __enter__(self):
        self.connect()
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        self.disconnect()
    
    def fetch_unread_emails(
        self,
        since_days: Optional[int] = None
    ) -> List[Dict[str, Any]]:
        """
        Fetch unread emails matching filter criteria.
        
        Returns:
            List of email dictionaries containing metadata and content
        """
        emails = []
        
        if not self.connection:
            self.logger.error("Not connected to email server")
            return emails
        
        try:
            # Select inbox
            status, _ = self.connection.select(self.settings.email_inbox_folder)
            if status != 'OK':
                self.logger.error(f"Failed to select inbox: {self.settings.email_inbox_folder}")
                return emails
            
            # Build search criteria
            search_criteria = ['UNSEEN']
            
            if since_days:
                since_date = (datetime.now() - timedelta(days=since_days)).strftime('%d-%b-%Y')
                search_criteria.append(f'SINCE {since_date}')
            
            search_query = ' '.join(search_criteria)
            self.logger.info(f"Searching with criteria: {search_query}")
            
            # Search for emails
            status, message_ids = self.connection.search(None, search_query)
            
            if status != 'OK' or not message_ids[0]:
                self.logger.info("No unread emails found")
                return emails
            
            email_ids = message_ids[0].split()
            self.logger.info(f"Found {len(email_ids)} unread emails")
            
            # Fetch each email
            for email_id in email_ids:
                try:
                    email_data = self._fetch_email(email_id)
                    if email_data:  # Filters removed - processing all emails
                        email_data['id'] = email_id.decode() if isinstance(email_id, bytes) else email_id
                        emails.append(email_data)
                except Exception as e:
                    self.logger.error(f"Error processing email {email_id}: {str(e)}")
            
            self.logger.info(f"Processing all {len(emails)} emails (no filters applied)")
            
        except Exception as e:
            self.logger.error(f"Error fetching emails: {str(e)}")
        
        return emails
    
    def _fetch_email(self, email_id: bytes) -> Optional[Dict[str, Any]]:
        """Fetch and parse a single email."""
        try:
            status, msg_data = self.connection.fetch(email_id, '(RFC822)')
            
            if status != 'OK':
                return None
            
            raw_email = msg_data[0][1]
            msg = email.message_from_bytes(raw_email)
            
            # Extract headers
            subject = self._decode_header(msg.get('Subject', ''))
            sender = self._decode_header(msg.get('From', ''))
            date_str = msg.get('Date', '')
            
            # Parse date
            email_date = None
            try:
                email_date = parsedate_to_datetime(date_str)
            except:
                email_date = datetime.now(timezone.utc)
            
            # Extract sender email
            sender_email = self._extract_email_address(sender)
            
            email_data = {
                'subject': subject,
                'sender': sender,
                'sender_email': sender_email,
                'date': email_date,
                'date_str': date_str,
                'body': '',
                'attachments': [],
                'raw_message': msg
            }
            
            # Parse body and attachments
            self._parse_message_parts(msg, email_data)
            
            return email_data
            
        except Exception as e:
            self.logger.error(f"Error fetching email {email_id}: {str(e)}")
            return None
    
    def _parse_message_parts(self, msg: Message, email_data: Dict):
        """Parse email message parts to extract body and attachments."""
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get('Content-Disposition', ''))
                
                # Check for attachment
                if 'attachment' in content_disposition:
                    filename = self._get_attachment_filename(part)
                    if filename:
                        email_data['attachments'].append({
                            'filename': filename,
                            'content_type': content_type,
                            'payload': part.get_payload(decode=True)
                        })
                elif content_type == 'text/plain' and not email_data['body']:
                    email_data['body'] = self._get_text_content(part)
                elif content_type == 'text/html' and not email_data['body']:
                    email_data['body'] = self._get_text_content(part)
        else:
            content_type = msg.get_content_type()
            if content_type in ['text/plain', 'text/html']:
                email_data['body'] = self._get_text_content(msg)
    
    def _get_attachment_filename(self, part: Message) -> Optional[str]:
        """Extract filename from attachment part."""
        filename = part.get_filename()
        if filename:
            return self._decode_header(filename)
        
        # Try to get name from Content-Type
        content_type = part.get_content_type()
        params = part.get_params()
        if params:
            for param, value in params:
                if param == 'name':
                    return self._decode_header(value)
        
        return None
    
    def _get_text_content(self, part: Message) -> str:
        """Extract text content from message part."""
        try:
            payload = part.get_payload(decode=True)
            charset = part.get_content_charset() or 'utf-8'
            return payload.decode(charset, errors='ignore')
        except:
            return ''
    
    def _decode_header(self, header: str) -> str:
        """Decode email header."""
        if not header:
            return ''
        
        decoded_parts = decode_header(header)
        result = []
        
        for part, charset in decoded_parts:
            if isinstance(part, bytes):
                charset = charset or 'utf-8'
                try:
                    result.append(part.decode(charset))
                except:
                    result.append(part.decode('utf-8', errors='ignore'))
            else:
                result.append(part)
        
        return ''.join(result)
    
    def _extract_email_address(self, from_header: str) -> str:
        """Extract email address from From header."""
        match = re.search(r'<([^>]+)>', from_header)
        if match:
            return match.group(1).lower()
        
        # If no angle brackets, assume the whole string is email
        return from_header.lower().strip()
    
    def _matches_filters(self, email_data: Dict[str, Any]) -> bool:
        """Check if email matches configured filters."""
        # Check sender filter
        if self.settings.allowed_senders:
            sender_email = email_data.get('sender_email', '').lower()
            if sender_email not in [s.lower() for s in self.settings.allowed_senders]:
                self.logger.debug(f"Email from {sender_email} not in allowed senders list")
                return False
        
        # Check subject keywords
        if self.settings.subject_keywords:
            subject = email_data.get('subject', '').lower()
            if not any(keyword.lower() in subject for keyword in self.settings.subject_keywords):
                self.logger.debug(f"Subject '{subject}' doesn't match any keywords")
                return False
        
        # Check date filter
        if self.settings.date_filter_days > 0:
            email_date = email_data.get('date')
            if email_date:
                cutoff_date = datetime.now(timezone.utc) - timedelta(days=self.settings.date_filter_days)
                if email_date < cutoff_date:
                    self.logger.debug(f"Email date {email_date} is older than cutoff")
                    return False
        
        return True
    
    def download_attachment(
        self,
        attachment: Dict[str, Any],
        download_folder: str
    ) -> Optional[str]:
        """
        Download attachment to local folder.
        
        Returns:
            Path to downloaded file or None if failed
        """
        try:
            folder = Path(download_folder)
            folder.mkdir(parents=True, exist_ok=True)
            
            original_filename = attachment['filename']
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            
            # Extract extension
            if '.' in original_filename:
                name, ext = original_filename.rsplit('.', 1)
                filename = f"{name}_{timestamp}.{ext}"
            else:
                filename = f"{original_filename}_{timestamp}"
            
            file_path = folder / filename
            
            with open(file_path, 'wb') as f:
                f.write(attachment['payload'])
            
            self.logger.info(f"Downloaded attachment: {file_path}")
            return str(file_path)
            
        except Exception as e:
            self.logger.error(f"Failed to download attachment: {str(e)}")
            return None
    
    def mark_as_read(self, email_id: str) -> bool:
        """Mark email as read."""
        try:
            self.connection.store(email_id.encode(), '+FLAGS', '\\Seen')
            self.logger.info(f"Marked email {email_id} as read")
            return True
        except Exception as e:
            self.logger.error(f"Failed to mark email as read: {str(e)}")
            return False
    
    def move_to_folder(self, email_id: str, target_folder: str) -> bool:
        """Move email to specified folder."""
        try:
            # Copy to target folder
            self.connection.copy(email_id.encode(), target_folder)
            # Mark as deleted in inbox
            self.connection.store(email_id.encode(), '+FLAGS', '\\Deleted')
            # Expunge
            self.connection.expunge()
            
            self.logger.info(f"Moved email {email_id} to folder {target_folder}")
            return True
            
        except Exception as e:
            self.logger.error(f"Failed to move email: {str(e)}")
            return False
    
    def has_excel_attachments(self, email_data: Dict[str, Any]) -> bool:
        """Check if email has Excel attachments matching configured extensions."""
        if not email_data.get('attachments'):
            return False
        
        extensions = [ext.lower() for ext in self.settings.file_extensions]
        
        for attachment in email_data['attachments']:
            filename = attachment['filename'].lower()
            if any(filename.endswith(ext) for ext in extensions):
                return True
        
        return False
