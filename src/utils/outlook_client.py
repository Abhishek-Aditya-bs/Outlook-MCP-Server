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
        """Comprehensive search in both subject and body using exact phrase matching."""
        emails = []
        
        try:
            # Build filter to search in both subject AND body
            # Using OR to find emails that contain the search text in either field
            filter_str = f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{search_text}%' OR \"urn:schemas:httpmail:textdescription\" LIKE '%{search_text}%'"
            
            logger.info(f"Searching {mailbox_type} mailbox for '{search_text}' in subject and body")
            
            # Get items from inbox
            items = inbox_folder.Items
            
            # Sort by ReceivedTime descending for recent emails first
            items.Sort("[ReceivedTime]", True)
            
            # Apply filter
            try:
                filtered_items = items.Restrict(filter_str)
                logger.info(f"Filter found {filtered_items.Count} potential matches")
            except:
                # Fallback to manual search if SQL filter fails
                logger.info("SQL filter failed, using fallback search")
                filtered_items = self._comprehensive_fallback_search(items, search_text, max_results)
            
            # Process filtered results
            count = 0
            for item in filtered_items:
                if count >= max_results:
                    break
                    
                try:
                    email_data = self._extract_email_data_optimized(
                        item, 
                        inbox_folder.Name, 
                        mailbox_type
                    )
                    if email_data:
                        emails.append(email_data)
                        count += 1
                except Exception as e:
                    logger.debug(f"Error processing email: {e}")
                    continue
            
            # Search other folders if enabled and we need more results
            if count < max_results and config.get_bool('search_all_folders', True):
                additional_emails = self._search_other_folders_comprehensive(
                    inbox_folder.Parent,
                    search_text,
                    mailbox_type,
                    max_results - count
                )
                emails.extend(additional_emails)
            
        except Exception as e:
            logger.error(f"Error in comprehensive search: {e}")
        
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
    
    def _search_other_folders_comprehensive(self, store, search_text: str,
                                           mailbox_type: str, max_results: int) -> List[Dict[str, Any]]:
        """Search other folders for the search text in both subject and body."""
        emails = []
        
        # Search key folders
        key_folders = ['Sent Items', 'Drafts']
        
        for folder_name in key_folders:
            if len(emails) >= max_results:
                break
                
            try:
                folder = self._get_folder_by_name(store, folder_name)
                if folder:
                    items = folder.Items
                    items.Sort("[ReceivedTime]", True)
                    
                    # Search in folder
                    count = 0
                    search_lower = search_text.lower()
                    
                    for item in items:
                        if count >= 10:  # Limit per folder
                            break
                        try:
                            subject = getattr(item, 'Subject', '').lower()
                            body = getattr(item, 'Body', '').lower()
                            
                            # Check both subject and body
                            if search_lower in subject or search_lower in body:
                                email_data = self._extract_email_data_optimized(
                                    item, folder_name, mailbox_type
                                )
                                if email_data:
                                    emails.append(email_data)
                                    count += 1
                        except:
                            continue
            except Exception as e:
                logger.debug(f"Error searching {folder_name}: {e}")
        
        return emails
    
    def _search_mailbox_optimized(self, inbox_folder, subject_pattern: str, 
                                  mailbox_type: str, max_results: int) -> List[Dict[str, Any]]:
        """Optimized search using Outlook's Restrict method."""
        emails = []
        
        try:
            # Build filter string for Outlook
            # Use Subject filter which is much faster than iterating
            filter_str = f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{subject_pattern}%'"
            
            # Alternative simpler filter (if SQL doesn't work)
            # filter_str = f"[Subject] = '{subject_pattern}'"
            
            logger.info(f"Searching {mailbox_type} mailbox with filter: {filter_str}")
            
            # Get items from inbox
            items = inbox_folder.Items
            
            # Sort by ReceivedTime descending for recent emails first
            items.Sort("[ReceivedTime]", True)
            
            # Apply filter - this is MUCH faster than iterating
            try:
                filtered_items = items.Restrict(filter_str)
                logger.info(f"Filter found {filtered_items.Count} potential matches")
            except:
                # Fallback to simpler search if SQL filter fails
                logger.info("SQL filter failed, trying simple search")
                filtered_items = self._simple_search(items, subject_pattern, max_results)
            
            # Process filtered results
            count = 0
            for item in filtered_items:
                if count >= max_results:
                    break
                    
                try:
                    email_data = self._extract_email_data_optimized(
                        item, 
                        inbox_folder.Name, 
                        mailbox_type
                    )
                    if email_data:
                        emails.append(email_data)
                        count += 1
                except Exception as e:
                    logger.debug(f"Error processing email: {e}")
                    continue
            
            # If we need more results and search_all_folders is enabled
            if count < max_results and config.get_bool('search_all_folders', True):
                # Search other folders
                additional_emails = self._search_other_folders_optimized(
                    inbox_folder.Parent,
                    subject_pattern,
                    mailbox_type,
                    max_results - count
                )
                emails.extend(additional_emails)
            
        except Exception as e:
            logger.error(f"Error in optimized search: {e}")
        
        return emails
    
    def _simple_search(self, items, pattern: str, max_results: int):
        """Simple fallback search method."""
        results = []
        pattern_lower = pattern.lower()
        count = 0
        
        for item in items:
            if count >= max_results:
                break
            try:
                subject = getattr(item, 'Subject', '').lower()
                if pattern_lower in subject:
                    results.append(item)
                    count += 1
            except:
                continue
        
        return results
    
    def _search_other_folders_optimized(self, store, subject_pattern: str, 
                                       mailbox_type: str, max_results: int) -> List[Dict[str, Any]]:
        """Search other folders if needed - optimized version."""
        emails = []
        
        # Only search key folders
        key_folders = ['Sent Items', 'Drafts']
        
        for folder_name in key_folders:
            if len(emails) >= max_results:
                break
                
            try:
                folder = self._get_folder_by_name(store, folder_name)
                if folder:
                    items = folder.Items
                    items.Sort("[ReceivedTime]", True)
                    
                    # Quick search in folder
                    count = 0
                    for item in items:
                        if count >= 10:  # Limit per folder
                            break
                        try:
                            subject = getattr(item, 'Subject', '').lower()
                            if subject_pattern.lower() in subject:
                                email_data = self._extract_email_data_optimized(
                                    item, folder_name, mailbox_type
                                )
                                if email_data:
                                    emails.append(email_data)
                                    count += 1
                        except:
                            continue
            except Exception as e:
                logger.debug(f"Error searching {folder_name}: {e}")
        
        return emails
    
    def _extract_email_data_optimized(self, item, folder_name: str, 
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
