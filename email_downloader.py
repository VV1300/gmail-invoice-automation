"""
Gmail Invoice Downloader for RPA System
Downloads invoice attachments from Gmail and saves them to input directory
"""
import os
import email
import imaplib
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from pathlib import Path
import re
from datetime import datetime, timedelta
from typing import List, Dict, Any, Optional
import time

from utils.logger import logger
from utils.exceptions import EmailError, NetworkError
from config import EMAIL_CONFIG, INPUT_DIR

class GmailDownloader:
    """Downloads invoice attachments from Gmail"""
    
    def __init__(self, email_config: Dict[str, Any] = None):
        self.config = email_config or EMAIL_CONFIG
        self.input_dir = Path(INPUT_DIR)
        self.input_dir.mkdir(exist_ok=True)
        
        # Invoice search patterns - focus on subject/title keywords
        self.invoice_keywords = [
            'invoice', 'bill', 'receipt', 'statement', 'payment',
            'INV-', 'INVOICE-', 'BILL-', 'RECEIPT-', 'PAYMENT-',
            'invoice #', 'bill #', 'receipt #', 'statement #',
            'invoice number', 'bill number', 'receipt number',
            'monthly invoice', 'monthly bill', 'monthly statement',
            'quarterly invoice', 'quarterly bill', 'quarterly statement',
            'annual invoice', 'annual bill', 'annual statement',
            'service invoice', 'service bill', 'service receipt',
            'consulting invoice', 'consulting bill',
            'development invoice', 'development bill',
            'software invoice', 'software bill',
            'maintenance invoice', 'maintenance bill',
            'support invoice', 'support bill'
        ]
        
        # Supported attachment extensions
        self.supported_extensions = [
            '.pdf', '.xlsx', '.xls', '.docx', '.doc'
        ]
    
    # def connect_to_gmail(self) -> imaplib.IMAP4_SSL:
    def connect_to_gmail(self):
        """Connect to Gmail using IMAP"""
        try:
            logger.info("Connecting to Gmail...")
            
            # Connect to Gmail IMAP server
            mail = imaplib.IMAP4_SSL('imap.gmail.com', 993)
            
            # Login with email and app password
            mail.login(self.config['sender_email'], self.config['sender_password'])
            
            logger.info("Successfully connected to Gmail")
            return mail
            
        except Exception as e:
            logger.error(f"Failed to connect to Gmail: {str(e)}")
            raise EmailError(f"Gmail connection failed: {str(e)}")
    
    # def search_invoices(self, mail: imaplib.IMAP4_SSL, days_back: int = 30) -> List[str]:
    def search_invoices(self, mail: imaplib.IMAP4_SSL, days_back: int = 30):
        """Search for emails containing invoices based on subject/title"""
        try:
            logger.info(f"Searching for invoice emails from last {days_back} days...")
            
            # Select inbox
            mail.select('INBOX')
            
            # Calculate date range
            date_since = (datetime.now() - timedelta(days=days_back)).strftime('%d-%b-%Y')
            
            # Search criteria
            search_criteria = f'(SINCE "{date_since}")'
            
            # Search for emails
            status, message_numbers = mail.search(None, search_criteria)
            
            if status != 'OK':
                raise EmailError("Failed to search emails")
            
            email_ids = message_numbers[0].split()
            invoice_emails = []
            
            logger.info(f"Found {len(email_ids)} emails to check")
            
            # Check each email for invoice content based on subject
            for email_id in email_ids:
                try:
                    # Fetch email
                    status, msg_data = mail.fetch(email_id, '(RFC822)')
                    
                    if status != 'OK':
                        continue
                    
                    email_body = msg_data[0][1]
                    email_message = email.message_from_bytes(email_body)
                    
                    # Get email subject (title)
                    subject = email_message.get('subject', '').lower()
                    
                    # Check if subject contains invoice keywords
                    is_invoice_email = any(keyword in subject for keyword in self.invoice_keywords)
                    
                    # Check for attachments as secondary criteria
                    has_attachments = self._has_supported_attachments(email_message)
                    
                    # Primary detection: Subject contains invoice keywords
                    # Secondary detection: Has supported attachments (in case subject is generic)
                    if is_invoice_email or has_attachments:
                        invoice_emails.append(email_id)
                        logger.info(f"Found invoice email: {email_message.get('subject', 'No Subject')}")
                        logger.info(f"  - Subject keywords detected: {is_invoice_email}")
                        logger.info(f"  - Has attachments: {has_attachments}")
                        
                except Exception as e:
                    logger.warning(f"Error processing email {email_id}: {str(e)}")
                    continue
            
            logger.info(f"Found {len(invoice_emails)} potential invoice emails")
            return invoice_emails
            
        except Exception as e:
            logger.error(f"Error searching for invoices: {str(e)}")
            raise EmailError(f"Failed to search invoices: {str(e)}")
    
    def _has_supported_attachments(self, email_message) -> bool:
        """Check if email has supported attachments"""
        try:
            for part in email_message.walk():
                if part.get_content_maintype() == 'multipart':
                    continue
                
                filename = part.get_filename()
                if filename:
                    file_ext = Path(filename).suffix.lower()
                    if file_ext in self.supported_extensions:
                        return True
            
            return False
            
        except Exception as e:
            logger.warning(f"Error checking attachments: {str(e)}")
            return False
    
    # def download_attachments(self, mail: imaplib.IMAP4_SSL, email_ids: List[str]) -> List[str]:
    def download_attachments(self, mail: imaplib.IMAP4_SSL, email_ids: List[str]):
        """Download attachments from invoice emails"""
        downloaded_files = []
        
        try:
            logger.info(f"Downloading attachments from {len(email_ids)} emails...")
            
            for email_id in email_ids:
                try:
                    # Fetch email
                    status, msg_data = mail.fetch(email_id, '(RFC822)')
                    
                    if status != 'OK':
                        continue
                    
                    email_body = msg_data[0][1]
                    email_message = email.message_from_bytes(email_body)
                    
                    # Get email metadata
                    subject = email_message.get('subject', 'Unknown')
                    from_email = email_message.get('from', 'Unknown')
                    date = email_message.get('date', 'Unknown')
                    
                    logger.info(f"Processing email: {subject}")
                    
                    # Download attachments
                    attachment_files = self._download_email_attachments(
                        email_message, subject, from_email, date
                    )
                    
                    downloaded_files.extend(attachment_files)
                    
                except Exception as e:
                    logger.error(f"Error downloading from email {email_id}: {str(e)}")
                    continue
            
            logger.info(f"Successfully downloaded {len(downloaded_files)} files")
            return downloaded_files
            
        except Exception as e:
            logger.error(f"Error downloading attachments: {str(e)}")
            raise EmailError(f"Failed to download attachments: {str(e)}")
    
    def _download_email_attachments(self, email_message, subject: str, from_email: str, date: str) -> List[str]:
        """Download attachments from a single email"""
        downloaded_files = []
        
        try:
            # Create a safe filename from subject
            safe_subject = re.sub(r'[^\w\s-]', '', subject)
            safe_subject = re.sub(r'[-\s]+', '-', safe_subject)
            safe_subject = safe_subject[:50]  # Limit length
            
            # Process each part of the email
            for part in email_message.walk():
                if part.get_content_maintype() == 'multipart':
                    continue
                
                filename = part.get_filename()
                if not filename:
                    continue
                
                # Check if it's a supported file type
                file_ext = Path(filename).suffix.lower()
                if file_ext not in self.supported_extensions:
                    logger.debug(f"Skipping unsupported file: {filename}")
                    continue
                
                # Create unique filename
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                base_name = Path(filename).stem
                new_filename = f"{safe_subject}_{base_name}_{timestamp}{file_ext}"
                file_path = self.input_dir / new_filename
                
                # Download attachment
                try:
                    with open(file_path, 'wb') as f:
                        f.write(part.get_payload(decode=True))
                    
                    downloaded_files.append(str(file_path))
                    logger.info(f"Downloaded: {new_filename}")
                    
                    # Add metadata file
                    self._create_metadata_file(file_path, subject, from_email, date)
                    
                except Exception as e:
                    logger.error(f"Error downloading {filename}: {str(e)}")
                    continue
            
            return downloaded_files
            
        except Exception as e:
            logger.error(f"Error processing email attachments: {str(e)}")
            return []
    
    def _create_metadata_file(self, file_path: Path, subject: str, from_email: str, date: str):
        """Create metadata file for downloaded attachment"""
        try:
            metadata_file = file_path.with_suffix('.metadata.txt')
            
            with open(metadata_file, 'w', encoding='utf-8') as f:
                f.write(f"Original Email Subject: {subject}\n")
                f.write(f"From: {from_email}\n")
                f.write(f"Date: {date}\n")
                f.write(f"Downloaded: {datetime.now().isoformat()}\n")
                f.write(f"File: {file_path.name}\n")
            
        except Exception as e:
            logger.warning(f"Error creating metadata file: {str(e)}")
    
    def mark_emails_as_read(self, mail: imaplib.IMAP4_SSL, email_ids: List[str]):
        """Mark processed emails as read"""
        try:
            logger.info("Marking processed emails as read...")
            
            for email_id in email_ids:
                try:
                    mail.store(email_id, '+FLAGS', '\\Seen')
                except Exception as e:
                    logger.warning(f"Error marking email {email_id} as read: {str(e)}")
            
            logger.info("Emails marked as read")
            
        except Exception as e:
            logger.error(f"Error marking emails as read: {str(e)}")
    
    def download_invoices(self, days_back: int = 30, mark_as_read: bool = True) -> Dict[str, Any]:
        """Main method to download invoices from Gmail"""
        try:
            logger.info("Starting Gmail invoice download process...")
            
            # Connect to Gmail
            mail = self.connect_to_gmail()
            
            try:
                # Search for invoice emails
                invoice_email_ids = self.search_invoices(mail, days_back)
                
                if not invoice_email_ids:
                    logger.info("No invoice emails found")
                    return {
                        'success': True,
                        'emails_found': 0,
                        'files_downloaded': 0,
                        'downloaded_files': []
                    }
                
                # Download attachments
                downloaded_files = self.download_attachments(mail, invoice_email_ids)
                
                # Mark emails as read if requested
                if mark_as_read:
                    self.mark_emails_as_read(mail, invoice_email_ids)
                
                # Disconnect from Gmail
                mail.close()
                mail.logout()
                
                logger.info("Gmail download process completed successfully")
                
                return {
                    'success': True,
                    'emails_found': len(invoice_email_ids),
                    'files_downloaded': len(downloaded_files),
                    'downloaded_files': downloaded_files
                }
                
            except Exception as e:
                mail.close()
                mail.logout()
                raise e
                
        except Exception as e:
            logger.error(f"Gmail download process failed: {str(e)}")
            return {
                'success': False,
                'error': str(e),
                'emails_found': 0,
                'files_downloaded': 0,
                'downloaded_files': []
            }
    
    def test_connection(self) -> bool:
        """Test Gmail connection"""
        try:
            logger.info("Testing Gmail connection...")
            
            mail = self.connect_to_gmail()
            mail.close()
            mail.logout()
            
            logger.info("Gmail connection test successful")
            return True
            
        except Exception as e:
            logger.error(f"Gmail connection test failed: {str(e)}")
            return False 