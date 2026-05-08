"""Outlook COM Automation Service - requires Microsoft Outlook desktop app."""

import win32com.client
import pythoncom
from pathlib import Path
from datetime import datetime, timedelta, timezone
from typing import List, Optional, Dict, Any

from config.settings import Settings
from utils.logger import get_logger


class OutlookCOMAutoService:
    """Email service using Outlook COM automation."""
    
    def __init__(self, settings: Settings):
        self.settings = settings
        self.logger = get_logger('OutlookCOMAutoService', settings.log_folder)
        self.outlook = None
        self.namespace = None
        self.inbox = None
    
    def connect(self) -> bool:
        """Connect to Outlook via COM."""
        try:
            self.logger.info("Connecting to Outlook via COM...")
            # Initialize COM for this thread
            pythoncom.CoInitialize()
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            self.inbox = self.namespace.GetDefaultFolder(6)  # 6 = inbox
            self.logger.info(f"Connected. Unread: {self.inbox.UnReadItemCount}")
            return True
        except Exception as e:
            self.logger.error(f"Failed to connect: {str(e)}")
            return False
    
    def disconnect(self):
        """Disconnect from Outlook."""
        self.outlook = None
        self.namespace = None
        self.inbox = None
        # Uninitialize COM
        try:
            pythoncom.CoUninitialize()
        except:
            pass
    
    def __enter__(self):
        self.connect()
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        self.disconnect()
    
    def fetch_unread_emails(self, since_days: Optional[int] = None) -> List[Dict[str, Any]]:
        """Fetch unread emails from Outlook."""
        emails = []
        if not self.inbox:
            self.logger.error("Not connected")
            return emails
        
        try:
            items = self.inbox.Items
            items.Sort("[ReceivedTime]", True)
            
            cutoff_date = None
            if since_days:
                cutoff_date = datetime.now(timezone.utc) - timedelta(days=since_days)
            
            for item in items:
                try:
                    if not item.UnRead:
                        continue
                    
                    if cutoff_date:
                        try:
                            if item.ReceivedTime.replace(tzinfo=timezone.utc) < cutoff_date:
                                continue
                        except:
                            pass
                    
                    email_data = self._convert_item(item)
                    if email_data:
                        emails.append(email_data)
                        
                except Exception as e:
                    continue
            
            self.logger.info(f"Found {len(emails)} unread emails")
            
        except Exception as e:
            self.logger.error(f"Error: {str(e)}")
        
        return emails
    
    def _convert_item(self, item) -> Optional[Dict[str, Any]]:
        """Convert Outlook item to standard format."""
        try:
            sender_email = ''
            try:
                if hasattr(item, 'Sender') and item.Sender:
                    sender_email = getattr(item.Sender, 'Address', '') or getattr(item.Sender, 'SMTPAddress', '')
            except:
                pass
            
            attachments = []
            if hasattr(item, 'Attachments'):
                for att in item.Attachments:
                    attachments.append({
                        'filename': att.FileName,
                        'outlook_attachment': att
                    })
            
            return {
                'id': item.EntryID,
                'subject': getattr(item, 'Subject', ''),
                'sender': f"{getattr(item, 'SenderName', '')} <{sender_email}>",
                'sender_email': sender_email.lower(),
                'date': getattr(item, 'ReceivedTime', datetime.now()),
                'attachments': attachments,
                'outlook_item': item
            }
        except Exception as e:
            return None
    
    def download_attachment(self, attachment: Dict[str, Any], download_folder: str) -> Optional[str]:
        """Download attachment."""
        try:
            folder = Path(download_folder)
            folder.mkdir(parents=True, exist_ok=True)
            
            original = attachment['filename']
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f"{original.rsplit('.', 1)[0]}_{timestamp}.{original.rsplit('.', 1)[1]}" if '.' in original else f"{original}_{timestamp}"
            
            # Use absolute path with forward slashes for COM compatibility
            file_path = folder.resolve() / filename
            file_path_str = str(file_path).replace('/', '\\')  # Windows backslash
            
            # Ensure parent exists
            file_path.parent.mkdir(parents=True, exist_ok=True)
            
            # Save attachment using COM
            outlook_att = attachment['outlook_attachment']
            outlook_att.SaveAsFile(file_path_str)
            
            self.logger.info(f"Downloaded: {file_path}")
            return str(file_path)
            
        except Exception as e:
            self.logger.error(f"Download failed: {str(e)}")
            return None
    
    def mark_as_read(self, email_id: str) -> bool:
        """Mark email as read."""
        try:
            # Find item by EntryID
            item = self.namespace.GetItemFromID(email_id)
            if item:
                item.UnRead = False
                self.logger.info(f"Marked {email_id} as read")
                return True
            return False
        except Exception as e:
            self.logger.error(f"Failed: {str(e)}")
            return False
    
    def move_to_folder(self, email_id: str, target_folder: str) -> bool:
        """Move email to folder."""
        try:
            # Find destination folder
            dest_folder = None
            for folder in self.inbox.Folders:
                if folder.Name == target_folder:
                    dest_folder = folder
                    break
            
            # Create folder if it doesn't exist
            if not dest_folder:
                self.logger.info(f"Creating folder: {target_folder}")
                dest_folder = self.inbox.Folders.Add(target_folder)
            
            # Move item
            item = self.namespace.GetItemFromID(email_id)
            if item:
                item.Move(dest_folder)
                self.logger.info(f"Moved {email_id} to {target_folder}")
                return True
            return False
            
        except Exception as e:
            self.logger.error(f"Move failed: {str(e)}")
            return False
    
    def has_excel_attachments(self, email_data: Dict[str, Any]) -> bool:
        """Check for Excel attachments."""
        if not email_data.get('attachments'):
            return False
        
        extensions = [ext.lower() for ext in self.settings.file_extensions]
        
        for att in email_data['attachments']:
            filename = att['filename'].lower()
            if any(filename.endswith(ext) for ext in extensions):
                return True
        
        return False
