"""Simple Outlook client for mailbox access and email search."""

import win32com.client
from datetime import datetime, timedelta
from typing import List, Dict, Any, Optional
import logging

from ..config.config_reader import config

logger = logging.getLogger(__name__)


class OutlookClient:
    """Simple client for accessing Outlook mailboxes."""
    
    def __init__(self):
        self.outlook = None
        self.namespace = None
        self.connected = False
    
    def connect(self) -> bool:
        """Connect to Outlook application."""
        try:
            logger.info("Connecting to Outlook...")
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            
            # Try Extended MAPI approach to potentially reduce security prompts (if enabled)
            if config.get_bool('use_extended_mapi_login', True):
                try:
                    logger.info("Attempting Extended MAPI login to reduce security prompts...")
                    # Parameters: Profile, Password, ShowDialog, NewSession
                    self.namespace.Logon(None, None, False, True)
                    logger.info("Extended MAPI login successful - may reduce security prompts")
                except Exception as logon_error:
                    logger.warning(f"Extended MAPI login failed: {logon_error}")
                    logger.info("Continuing with standard connection - security prompts may appear")
            else:
                logger.info("Extended MAPI login disabled in configuration")
            
            self.connected = True
            logger.info("Successfully connected to Outlook")
            return True
        except Exception as e:
            logger.error(f"Failed to connect to Outlook: {e}")
            self.connected = False
            return False
    
    def check_access(self) -> Dict[str, Any]:
        """Check access to personal and shared mailboxes."""
        if not self.connected:
            if not self.connect():
                return {"error": "Could not connect to Outlook"}
        
        result = {
            "outlook_connected": True,
            "personal_accessible": False,
            "shared_accessible": False,
            "shared_configured": bool(config.get('shared_mailbox_email')),
            "retention_personal_months": config.get_int('personal_retention_months', 6),
            "retention_shared_months": config.get_int('shared_retention_months', 12),
            "errors": []
        }
        
        # Test personal mailbox
        try:
            personal_inbox = self.namespace.GetDefaultFolder(6)  # 6 = Inbox
            if personal_inbox:
                result["personal_accessible"] = True
                result["personal_name"] = personal_inbox.Parent.DisplayName
        except Exception as e:
            result["errors"].append(f"Personal mailbox error: {str(e)}")
        
        # Test shared mailbox
        shared_email = config.get('shared_mailbox_email')
        if shared_email:
            try:
                shared_recipient = self.namespace.CreateRecipient(shared_email)
                shared_recipient.Resolve()
                
                if shared_recipient.Resolved:
                    shared_inbox = self.namespace.GetSharedDefaultFolder(shared_recipient, 6)
                    if shared_inbox:
                        result["shared_accessible"] = True
                        result["shared_name"] = shared_inbox.Parent.DisplayName
            except Exception as e:
                result["errors"].append(f"Shared mailbox error: {str(e)}")
        
        return result
    
    def search_emails_by_subject(self, subject: str, 
                                include_personal: bool = True, include_shared: bool = True) -> List[Dict[str, Any]]:
        """Search for emails by subject pattern."""
        if not self.connected:
            if not self.connect():
                return []
        
        all_emails = []
        
        # Search personal mailbox
        if include_personal:
            try:
                personal_store = self.namespace.GetDefaultFolder(6).Parent  # Get store from inbox
                personal_emails = self._search_store(personal_store, subject, 'personal')
                all_emails.extend(personal_emails)
                logger.info(f"Found {len(personal_emails)} emails in personal mailbox")
            except Exception as e:
                logger.error(f"Error searching personal mailbox: {e}")
        
        # Search shared mailbox
        if include_shared and config.get('shared_mailbox_email'):
            try:
                shared_email = config.get('shared_mailbox_email')
                shared_recipient = self.namespace.CreateRecipient(shared_email)
                shared_recipient.Resolve()
                
                if shared_recipient.Resolved:
                    shared_inbox = self.namespace.GetSharedDefaultFolder(shared_recipient, 6)
                    shared_store = shared_inbox.Parent
                    shared_emails = self._search_store(shared_store, subject, 'shared')
                    all_emails.extend(shared_emails)
                    logger.info(f"Found {len(shared_emails)} emails in shared mailbox")
            except Exception as e:
                logger.error(f"Error searching shared mailbox: {e}")
        
        # Sort by received time (newest first)
        all_emails.sort(key=lambda x: x.get('received_time', datetime.min), reverse=True)
        
        # Limit results
        max_results = config.get_int('max_search_results', 500)
        return all_emails[:max_results]
    
    def search_alerts(self, alert_pattern: str,
                     include_personal: bool = True, include_shared: bool = True) -> List[Dict[str, Any]]:
        """Search for production alerts using the exact pattern provided."""
        if not self.connected:
            if not self.connect():
                return []
        
        # Search for the exact alert pattern provided
        search_patterns = [alert_pattern]
        
        all_alerts = []
        
        # Search personal mailbox
        if include_personal:
            try:
                personal_store = self.namespace.GetDefaultFolder(6).Parent
                personal_alerts = self._search_store_for_patterns(personal_store, search_patterns, 'personal')
                all_alerts.extend(personal_alerts)
                logger.info(f"Found {len(personal_alerts)} alerts in personal mailbox")
            except Exception as e:
                logger.error(f"Error searching personal mailbox for alerts: {e}")
        
        # Search shared mailbox
        if include_shared and config.get('shared_mailbox_email'):
            try:
                shared_email = config.get('shared_mailbox_email')
                shared_recipient = self.namespace.CreateRecipient(shared_email)
                shared_recipient.Resolve()
                
                if shared_recipient.Resolved:
                    shared_inbox = self.namespace.GetSharedDefaultFolder(shared_recipient, 6)
                    shared_store = shared_inbox.Parent
                    shared_alerts = self._search_store_for_patterns(shared_store, search_patterns, 'shared')
                    all_alerts.extend(shared_alerts)
                    logger.info(f"Found {len(shared_alerts)} alerts in shared mailbox")
            except Exception as e:
                logger.error(f"Error searching shared mailbox for alerts: {e}")
        
        # Sort by received time (newest first)
        all_alerts.sort(key=lambda x: x.get('received_time', datetime.min), reverse=True)
        
        # Limit results
        max_results = config.get_int('max_search_results', 500)
        return all_alerts[:max_results]
    
    def _search_store(self, store, subject_pattern: str, mailbox_type: str) -> List[Dict[str, Any]]:
        """Search a specific mailbox store for emails by subject."""
        emails = []
        
        try:
            # Get all folders if configured to search all folders
            if config.get_bool('search_all_folders', True):
                folders = self._get_all_folders(store)
            else:
                # Just search inbox
                inbox = None
                for folder in store.GetRootFolder().Folders:
                    if folder.Name.lower() in ['inbox', 'posteingang']:
                        inbox = folder
                        break
                folders = [inbox] if inbox else []
            
            # Add sent items if configured
            if config.get_bool('include_sent_items', True):
                try:
                    sent_folder = self._find_folder_by_name(store, ['sent items', 'gesendete objekte', 'sent'])
                    if sent_folder:
                        folders.append(sent_folder)
                except Exception:
                    pass
            
            # Search each folder
            for folder in folders:
                if folder:
                    folder_emails = self._search_folder(folder, subject_pattern, mailbox_type)
                    emails.extend(folder_emails)
            
        except Exception as e:
            logger.error(f"Error searching {mailbox_type} store: {e}")
        
        return emails
    
    def _search_store_for_patterns(self, store, patterns: List[str], mailbox_type: str) -> List[Dict[str, Any]]:
        """Search store for multiple patterns (for alerts)."""
        emails = []
        
        try:
            # Get folders to search
            if config.get_bool('search_all_folders', True):
                folders = self._get_all_folders(store)
            else:
                inbox = None
                for folder in store.GetRootFolder().Folders:
                    if folder.Name.lower() in ['inbox', 'posteingang']:
                        inbox = folder
                        break
                folders = [inbox] if inbox else []
            
            # Search each folder for any matching pattern
            for folder in folders:
                if folder:
                    for pattern in patterns:
                        folder_emails = self._search_folder(folder, pattern, mailbox_type)
                        emails.extend(folder_emails)
            
            # Remove duplicates (emails matching multiple patterns)
            seen_ids = set()
            unique_emails = []
            for email in emails:
                email_id = email.get('entry_id', '')
                if email_id and email_id not in seen_ids:
                    seen_ids.add(email_id)
                    unique_emails.append(email)
                elif not email_id:  # Fallback for emails without entry_id
                    unique_emails.append(email)
            
            emails = unique_emails
            
        except Exception as e:
            logger.error(f"Error searching {mailbox_type} store for patterns: {e}")
        
        return emails
    
    def _search_folder(self, folder, pattern: str, mailbox_type: str) -> List[Dict[str, Any]]:
        """Search a specific folder for emails matching pattern."""
        emails = []
        
        try:
            items = folder.Items
            
            # Search through all items (entire mailbox like Outlook native search)
            count = 0
            batch_size = config.get_int('batch_processing_size', 50)
            
            for item in items:
                try:
                    subject = getattr(item, 'Subject', '').lower()
                    body = getattr(item, 'Body', '').lower()[:500]  # First 500 chars for performance
                    
                    # Check if pattern matches
                    if pattern.lower() in subject or pattern.lower() in body:
                        email_data = self._extract_email_data(item, folder.Name, mailbox_type)
                        if email_data:
                            emails.append(email_data)
                            count += 1
                    
                    # Process in batches for performance
                    if count >= batch_size:
                        break
                        
                except Exception as e:
                    logger.debug(f"Error processing email in folder {folder.Name}: {e}")
                    continue
            
        except Exception as e:
            logger.error(f"Error searching folder {getattr(folder, 'Name', 'Unknown')}: {e}")
        
        return emails
    
    def _extract_email_data(self, item, folder_name: str, mailbox_type: str) -> Dict[str, Any]:
        """Extract data from an email item."""
        try:
            # Get email properties
            subject = getattr(item, 'Subject', 'No Subject')
            sender_name = getattr(item, 'SenderName', 'Unknown')
            sender_email = getattr(item, 'SenderEmailAddress', '')
            received_time = getattr(item, 'ReceivedTime', datetime.now())
            body = getattr(item, 'Body', '')
            
            # Limit body if configured (0 means no limit)
            max_body_chars = config.get_int('max_body_chars', 0)
            if max_body_chars > 0 and len(body) > max_body_chars:
                body = body[:max_body_chars] + " [truncated]"
            
            # Clean HTML if configured
            if config.get_bool('clean_html_content', True):
                body = self._clean_html(body)
            
            # Get recipients
            recipients = []
            try:
                for recipient in item.Recipients:
                    recipients.append(getattr(recipient, 'Name', getattr(recipient, 'Address', '')))
            except Exception:
                pass
            
            return {
                'subject': subject,
                'sender_name': sender_name,
                'sender_email': sender_email,
                'recipients': recipients,
                'received_time': received_time,
                'body': body,
                'folder_name': folder_name,
                'mailbox_type': mailbox_type,
                'importance': getattr(item, 'Importance', 1),
                'size': getattr(item, 'Size', 0),
                'attachments_count': getattr(item.Attachments, 'Count', 0) if hasattr(item, 'Attachments') else 0,
                'unread': getattr(item, 'Unread', False),
                'entry_id': getattr(item, 'EntryID', '')
            }
            
        except Exception as e:
            logger.debug(f"Error extracting email data: {e}")
            return None
    
    def _get_all_folders(self, store) -> List:
        """Get all folders in a store recursively."""
        folders = []
        try:
            root_folder = store.GetRootFolder()
            self._traverse_folders(root_folder, folders)
        except Exception as e:
            logger.error(f"Error getting all folders: {e}")
        return folders
    
    def _traverse_folders(self, folder, folder_list: List):
        """Recursively traverse folder structure."""
        try:
            # Skip root folder itself
            if folder.Name not in ['Root - Mailbox', 'Mailbox']:
                folder_name = folder.Name.lower()
                
                # Skip system folders if not configured to include them
                if not config.get_bool('include_deleted_items', False):
                    if folder_name in ['deleted items', 'junk email', 'sync issues']:
                        return
                
                folder_list.append(folder)
            
            # Traverse subfolders
            try:
                for subfolder in folder.Folders:
                    self._traverse_folders(subfolder, folder_list)
            except Exception:
                pass
                
        except Exception as e:
            logger.debug(f"Error traversing folder: {e}")
    
    def _find_folder_by_name(self, store, names: List[str]):
        """Find folder by name (supports multiple possible names)."""
        try:
            root_folder = store.GetRootFolder()
            for folder in root_folder.Folders:
                if folder.Name.lower() in [name.lower() for name in names]:
                    return folder
        except Exception:
            pass
        return None
    
    def _clean_html(self, text: str) -> str:
        """Simple HTML cleaning."""
        import re
        
        # Remove HTML tags
        text = re.sub(r'<[^>]+>', '', text)
        
        # Decode common HTML entities
        html_entities = {
            '&amp;': '&',
            '&lt;': '<',
            '&gt;': '>',
            '&quot;': '"',
            '&#39;': "'",
            '&nbsp;': ' '
        }
        
        for entity, char in html_entities.items():
            text = text.replace(entity, char)
        
        # Clean up whitespace
        text = re.sub(r'\s+', ' ', text).strip()
        
        return text


# Global client instance
outlook_client = OutlookClient()
