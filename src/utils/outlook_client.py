"""High-performance Outlook client for mailbox access and email search."""

import win32com.client
from datetime import datetime, timedelta
from typing import List, Dict, Any
import logging
import pythoncom
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
        """Optimized search using Outlook's SQL filters for better performance."""
        emails = []
        found_ids = set()  # Track found emails to avoid duplicates
        
        # Escape special characters for SQL
        search_text_escaped = search_text.replace("'", "''").replace('"', '""')
        
        items = inbox_folder.Items
        items.Sort("[ReceivedTime]", True)
        
        # Build optimized filters to try in order of performance
        filters_to_try = [
            # 1. Subject search (fastest - always indexed)
            (f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{search_text_escaped}%'", "subject"),
            
            # 2. Try combined subject OR body with message body property
            (f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{search_text_escaped}%' OR "
             f"\"http://schemas.microsoft.com/mapi/proptag/0x1000001E\" LIKE '%{search_text_escaped}%'", "subject_or_body"),
            
            # 3. Try text description (often faster than full body)
            (f"@SQL=\"urn:schemas:httpmail:textdescription\" LIKE '%{search_text_escaped}%'", "text_description"),
        ]
        
        for filter_sql, filter_type in filters_to_try:
            try:
                logger.info(f"Trying {filter_type} filter for '{search_text}'")
                filtered_items = items.Restrict(filter_sql)
                count = filtered_items.Count
                
                if count > 0:
                    logger.info(f"Found {count} potential matches using {filter_type} filter")
                    
                    for item in filtered_items:
                        if len(emails) >= max_results:
                            break
                        
                        try:
                            entry_id = getattr(item, 'EntryID', '')
                            if entry_id and entry_id not in found_ids:
                                # For subject-only filter, verify the match
                                if filter_type == "subject":
                                    # Already matched in subject, just add it
                                    email_data = self._extract_email_data(
                                        item, inbox_folder.Name, mailbox_type
                                    )
                                    if email_data:
                                        emails.append(email_data)
                                        found_ids.add(entry_id)
                                else:
                                    # For body filters, verify it's actually in the content
                                    search_lower = search_text.lower()
                                    subject = getattr(item, 'Subject', '').lower()
                                    body = getattr(item, 'Body', '').lower()
                                    
                                    if search_lower in subject or search_lower in body:
                                        email_data = self._extract_email_data(
                                            item, inbox_folder.Name, mailbox_type
                                        )
                                        if email_data:
                                            emails.append(email_data)
                                            found_ids.add(entry_id)
                        except Exception as e:
                            logger.debug(f"Error processing item: {e}")
                            continue
                    
                    # If we got good results from this filter, we can stop
                    if len(emails) >= max_results * 0.5:  # Got at least 50% of what we need
                        logger.info(f"Got sufficient results from {filter_type} filter")
                        break
                        
            except Exception as e:
                logger.debug(f"Filter {filter_type} failed: {e}")
                continue
        
        # If we still need more results and no filter worked well, fall back to date-based search
        if len(emails) < max_results:
            logger.info(f"Need more results, using date-based search fallback")
            # Use smaller date ranges for faster performance
            date_ranges = [7, 30, 90, 180]  # Simplified ranges
            
            # Get max retention from config
            max_retention_days = config.get_int(
                f'{mailbox_type}_retention_months',
                config.get_int('personal_retention_months', 6) if mailbox_type == 'personal' else config.get_int('shared_retention_months', 12)
            ) * 30
            
            search_lower = search_text.lower()
            total_checked = 0
            
            for days in date_ranges:
                if len(emails) >= max_results:
                    break
                
                # Don't search beyond retention period
                if days > max_retention_days:
                    days = max_retention_days
                
                # Calculate date range
                from datetime import timedelta
                start_date = datetime.now() - timedelta(days=days)
                end_date = datetime.now() - timedelta(days=date_ranges[date_ranges.index(days)-1] if date_ranges.index(days) > 0 else 0)
                
                # Build date filter for this range
                if date_ranges.index(days) == 0:
                    # First range: just last N days
                    date_filter = f"[ReceivedTime] >= '{start_date.strftime('%m/%d/%Y')}'"
                else:
                    # Subsequent ranges: between dates to avoid re-checking
                    date_filter = f"[ReceivedTime] >= '{start_date.strftime('%m/%d/%Y')}' AND [ReceivedTime] < '{end_date.strftime('%m/%d/%Y')}'"
                
                try:
                    filtered_items = items.Restrict(date_filter)
                    range_count = filtered_items.Count
                    
                    if range_count == 0:
                        continue
                    
                    logger.info(f"Searching {range_count} emails from last {days} days")
                    
                    # Sort by date (newest first)
                    filtered_items.Sort("[ReceivedTime]", True)
                    
                    items_checked = 0
                    # Limit checks per range for better performance
                    max_per_range = min(range_count, 500)  # Reduced limit for faster response
                    
                    for item in filtered_items:
                        if len(emails) >= max_results or items_checked >= max_per_range:
                            break
                        
                        items_checked += 1
                        total_checked += 1
                        
                        try:
                            entry_id = getattr(item, 'EntryID', '')
                            if entry_id and entry_id not in found_ids:
                                # Quick subject check first
                                subject = getattr(item, 'Subject', '').lower()
                                body = getattr(item, 'Body', '').lower()
                                
                                # Check if search text is in subject or body
                                if search_lower in subject or search_lower in body:
                                    email_data = self._extract_email_data(
                                        item, inbox_folder.Name, mailbox_type
                                    )
                                    if email_data:
                                        emails.append(email_data)
                                        found_ids.add(entry_id)
                        except Exception as e:
                            logger.debug(f"Error checking email: {e}")
                            continue
                    
                    logger.info(f"Checked {items_checked} items in {days}-day range, found {len(emails)} total matches")
                    
                    # If we found enough matches, stop searching
                    if len(emails) >= max_results * 0.7:  # Found at least 70% of what we need
                        logger.info(f"Found sufficient matches, stopping search")
                        break
                        
                except Exception as e:
                    logger.warning(f"Date filter for {days} days failed: {e}")
                    continue
            
            logger.info(f"Date-based search checked {total_checked} items total, found {len(emails)} matches")
            
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
