"""Quick test for Outlook IMAP connection"""
import imaplib
import os
from config import Settings

settings = Settings.from_env()

print(f"Host: {settings.email_host}")
print(f"Port: {settings.email_port}")
print(f"User: {settings.email_username}")
print(f"Pass: {'*' * len(settings.email_password) if settings.email_password else 'EMPTY'}")
print(f"Provider: {settings.email_provider}")
print()

try:
    conn = imaplib.IMAP4_SSL(settings.email_host, settings.email_port)
    print("✅ Connected to server")
    
    # Try login
    conn.login(settings.email_username, settings.email_password)
    print("✅ Login successful")
    
    # List folders
    status, folders = conn.list()
    print(f"✅ Found {len(folders)} folders")
    
    # Select inbox
    conn.select('INBOX')
    print("✅ Inbox selected")
    
    # Check unread
    status, messages = conn.search(None, 'UNSEEN')
    print(f"✅ Unread emails: {len(messages[0].split())}")
    
    conn.logout()
    print("\n✅ All tests passed! Outlook IMAP is working.")
    
except Exception as e:
    print(f"\n❌ Error: {e}")
    print("\nPossible causes:")
    print("1. IMAP not enabled in Outlook settings")
    print("2. Security defaults blocking basic auth")
    print("3. Wrong app password format")
    print("4. Account requires OAuth2 (work/school account)")
