"""High-performance Outlook client for mailbox access and email search."""

import win32com.client
from datetime import datetime, timedelta
from typing import List, Dict, Any, Optional
import logging
import pythoncom
import threading
import re

from ..config.config_reader import config

logger = logging.getLogger(__name__)


class OutlookClient:
    """High-performance client for accessing Outlook mailboxes with optimized search."""
    
    def __init__(self):
        self.outlook = None
        self.namespace = None
        self.connected = False
        self._search_cache = {}  # Cache for search results
        self._folder_cache = {}  # Cache for folder references
    
    def connect(self) -> bool:
        """Connect to Outlook application."""
        try:
            logger.info("Connecting to Outlook...")
            
            # Initialize COM for thread
            pythoncom.CoInitialize()
            
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            
            # Try Extended MAPI approach to potentially reduce security prompts (if enabled)
            if config.get_bool('use_extended_mapi_login', True):
                try:
                    logger.info("Attempting Extended MAPI login to reduce security prompts...")
                    self.namespace.Logon(None, None, False, True)
                    logger.info("Extended MAPI login successful")
                except Exception as logon_error:
                    logger.warning(f"Extended MAPI login failed: {logon_error}")
            
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
                result["personal_name"] = self._get_store_display_name(personal_inbox)
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
                        result["shared_name"] = self._get_store_display_name(shared_inbox)
            except Exception as e:
                result["errors"].append(f"Shared mailbox error: {str(e)}")
        
        return result
    
    def search_emails(self, search_text: str, 
                     include_personal: bool = True, 
                     include_shared: bool = True) -> List[Dict[str, Any]]:
        """Search emails in both subject and body using exact phrase matching."""
        if not self.connected:
            if not self.connect():
                return []
        
        # Check cache first
        cache_key = f"{search_text}_{include_personal}_{include_shared}"
        if cache_key in self._search_cache:
            logger.info(f"Returning cached results for '{search_text}'")
            return self._search_cache[cache_key]
        
        all_emails = []
        max_results = config.get_int('max_search_results', 500)
        
        # Search personal mailbox
        if include_personal:
            personal_emails = self._search_mailbox_comprehensive(
                self.namespace.GetDefaultFolder(6), 
                search_text, 
                'personal',
                max_results
            )
            all_emails.extend(personal_emails)
            logger.info(f"Found {len(personal_emails)} emails in personal mailbox")
        
        # Search shared mailbox
        if include_shared and config.get('shared_mailbox_email'):
            try:
                shared_email = config.get('shared_mailbox_email')
                shared_recipient = self.namespace.CreateRecipient(shared_email)
                shared_recipient.Resolve()
                
                if shared_recipient.Resolved:
                    shared_inbox = self.namespace.GetSharedDefaultFolder(shared_recipient, 6)
                    shared_emails = self._search_mailbox_comprehensive(
                        shared_inbox,
                        search_text,
                        'shared',
                        max_results - len(all_emails)
                    )
                    all_emails.extend(shared_emails)
                    logger.info(f"Found {len(shared_emails)} emails in shared mailbox")
            except Exception as e:
                logger.error(f"Error searching shared mailbox: {e}")
        
        # Sort by received time (newest first)
        all_emails.sort(key=lambda x: x.get('received_time', datetime.min), reverse=True)
        
        # Cache results
        self._search_cache[cache_key] = all_emails[:max_results]
        
        return all_emails[:max_results]
    
    def search_emails_by_subject(self, subject: str, 
                                include_personal: bool = True, 
                                include_shared: bool = True) -> List[Dict[str, Any]]:
        """Legacy method - redirects to search_emails for backward compatibility."""
        return self.search_emails(subject, include_personal, include_shared)
    
    def _search_mailbox_comprehensive(self, inbox_folder, search_text: str,
                                      mailbox_type: str, max_results: int) -> List[Dict[str, Any]]:
        """Two-pass search strategy: first subject (fast), then body if needed."""
        emails = []
        found_ids = set()  # Track found emails to avoid duplicates
        
        try:
            # PASS 1: Search in SUBJECT first (fastest)
            subject_filter = f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{search_text}%'"
            logger.info(f"Pass 1: Searching {mailbox_type} mailbox subjects for '{search_text}'")
            
            items = inbox_folder.Items
            items.Sort("[ReceivedTime]", True)
            
            # Search subjects
            try:
                subject_items = items.Restrict(subject_filter)
                logger.info(f"Found {subject_items.Count} matches in subjects")
                
                # Process subject matches
                for item in subject_items:
                    if len(emails) >= max_results:
                        break
                    try:
                        entry_id = getattr(item, 'EntryID', '')
                        if entry_id and entry_id not in found_ids:
                            email_data = self._extract_email_data(
                                item, inbox_folder.Name, mailbox_type
                            )
                            if email_data:
                                emails.append(email_data)
                                found_ids.add(entry_id)
                    except Exception as e:
                        logger.debug(f"Error processing email: {e}")
                        continue
                        
            except Exception as e:
                logger.warning(f"Subject search failed: {e}")
            
            # PASS 2: Search in BODY if we need more results
            if len(emails) < max_results:
                logger.info(f"Pass 2: Searching email bodies for additional matches")
                
                # For body search, we need to iterate (slower but necessary)
                # Limit iteration to recent emails for performance
                search_lower = search_text.lower()
                items_checked = 0
                max_items_to_check = min(1000, max_results * 10)  # Limit iterations
                
                for item in items:
                    if len(emails) >= max_results or items_checked >= max_items_to_check:
                        break
                    
                    items_checked += 1
                    
                    try:
                        entry_id = getattr(item, 'EntryID', '')
                        if entry_id and entry_id not in found_ids:
                            # Quick subject check first (already in memory)
                            subject = getattr(item, 'Subject', '').lower()
                            if search_lower not in subject:  # Only check body if not in subject
                                body = getattr(item, 'Body', '').lower()
                                if search_lower in body:
                                    email_data = self._extract_email_data(
                                        item, inbox_folder.Name, mailbox_type
                                    )
                                    if email_data:
                                        emails.append(email_data)
                                        found_ids.add(entry_id)
                    except Exception as e:
                        logger.debug(f"Error checking email body: {e}")
                        continue
                
                logger.info(f"Checked {items_checked} items for body matches, found {len(emails)} total")
            
            # Search other folders if enabled and we need more results
            if len(emails) < max_results and config.get_bool('search_all_folders', True):
                additional_emails = self._search_other_folders(
                    inbox_folder.Parent,
                    search_text,
                    mailbox_type,
                    max_results - len(emails),
                    found_ids
                )
                emails.extend(additional_emails)
            
        except Exception as e:
            logger.error(f"Error in search: {e}")
        
        return emails
    
    def _comprehensive_fallback_search(self, items, search_text: str, max_results: int):
        """Fallback search that checks both subject and body."""
        results = []
        search_lower = search_text.lower()
        count = 0
        
        for item in items:
            if count >= max_results:
                break
            try:
                subject = getattr(item, 'Subject', '').lower()
                body = getattr(item, 'Body', '').lower()
                
                # Check if search text exists in either subject or body
                if search_lower in subject or search_lower in body:
                    results.append(item)
                    count += 1
            except:
                continue
        
        return results
    
    def _search_other_folders(self, store, search_text: str, mailbox_type: str, 
                             max_results: int, found_ids: set) -> List[Dict[str, Any]]:
        """Search other folders using two-pass strategy."""
        emails = []
        key_folders = ['Sent Items', 'Drafts']
        
        for folder_name in key_folders:
            if len(emails) >= max_results:
                break
                
            try:
                folder = self._get_folder_by_name(store, folder_name)
                if folder:
                    items = folder.Items
                    items.Sort("[ReceivedTime]", True)
                    
                    # Quick subject search first
                    subject_filter = f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{search_text}%'"
                    try:
                        filtered_items = items.Restrict(subject_filter)
                        for item in filtered_items:
                            if len(emails) >= max_results:
                                break
                            entry_id = getattr(item, 'EntryID', '')
                            if entry_id and entry_id not in found_ids:
                                email_data = self._extract_email_data(
                                    item, folder_name, mailbox_type
                                )
                                if email_data:
                                    emails.append(email_data)
                                    found_ids.add(entry_id)
                    except:
                        # Fallback to limited iteration if filter fails
                        count = 0
                        search_lower = search_text.lower()
                        for item in items:
                            if count >= 10 or len(emails) >= max_results:
                                break
                            entry_id = getattr(item, 'EntryID', '')
                            if entry_id and entry_id not in found_ids:
                                subject = getattr(item, 'Subject', '').lower()
                                if search_lower in subject:
                                    email_data = self._extract_email_data(
                                        item, folder_name, mailbox_type
                                    )
                                    if email_data:
                                        emails.append(email_data)
                                        found_ids.add(entry_id)
                                        count += 1
            except Exception as e:
                logger.debug(f"Error searching {folder_name}: {e}")
        
        return emails
    
    def _extract_email_data(self, item, folder_name: str, 
                           mailbox_type: str) -> Dict[str, Any]:
        """Extract complete email data including full body."""
        try:
            # Get the full email body
            body = getattr(item, 'Body', '')
            
            # Apply max_body_chars if configured (0 means no limit)
            max_body_chars = config.get_int('max_body_chars', 0)
            if max_body_chars > 0 and len(body) > max_body_chars:
                body = body[:max_body_chars] + " [truncated]"
            
            # Clean HTML if configured
            if config.get_bool('clean_html_content', True) and body:
                body = self._clean_html(body)
            
            # Get recipients list
            recipients = []
            try:
                for recipient in item.Recipients:
                    recipients.append(getattr(recipient, 'Name', getattr(recipient, 'Address', '')))
            except:
                pass
            
            return {
                'subject': getattr(item, 'Subject', 'No Subject'),
                'sender_name': getattr(item, 'SenderName', 'Unknown'),
                'sender_email': getattr(item, 'SenderEmailAddress', ''),
                'recipients': recipients,
                'received_time': getattr(item, 'ReceivedTime', datetime.now()),
                'folder_name': folder_name,
                'mailbox_type': mailbox_type,
                'importance': getattr(item, 'Importance', 1),
                'body': body,  # Full body for summarization
                'size': getattr(item, 'Size', 0),
                'attachments_count': getattr(item.Attachments, 'Count', 0) if hasattr(item, 'Attachments') else 0,
                'unread': getattr(item, 'Unread', False),
                'entry_id': getattr(item, 'EntryID', '')
            }
        except Exception as e:
            logger.error(f"Error extracting email data: {e}")
            return None
    
    def _get_store_display_name(self, folder) -> str:
        """Safely get store display name from a folder."""
        try:
            if hasattr(folder, 'Parent'):
                parent = folder.Parent
                if hasattr(parent, 'DisplayName'):
                    return parent.DisplayName
                elif hasattr(parent, 'Name'):
                    return parent.Name
            return "Mailbox"
        except:
            return "Mailbox"
    
    def _get_folder_by_name(self, store, name: str):
        """Get folder by name from cache or store."""
        cache_key = f"{id(store)}_{name}"
        
        if cache_key in self._folder_cache:
            return self._folder_cache[cache_key]
        
        try:
            for folder in store.GetRootFolder().Folders:
                if folder.Name.lower() == name.lower():
                    self._folder_cache[cache_key] = folder
                    return folder
        except:
            pass
        
        return None
    
    def clear_cache(self):
        """Clear search cache."""
        self._search_cache.clear()
        self._folder_cache.clear()
        logger.info("Cache cleared")
    
    def _clean_html(self, text: str) -> str:
        """Clean HTML from email body."""
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
