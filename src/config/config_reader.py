"""Simple configuration reader for properties file."""

import os
from typing import Dict, Any, List


class ConfigReader:
    """Reads configuration from config.properties file."""
    
    def __init__(self, config_file: str = "config.properties"):
        self.config_file = config_file
        self.config = {}
        self.load_config()
    
    def load_config(self):
        """Load configuration from properties file."""
        # Look for config file in the config directory
        config_path = os.path.join(os.path.dirname(__file__), self.config_file)
        
        if not os.path.exists(config_path):
            print(f"Warning: Config file {config_path} not found. Using defaults.")
            self._set_defaults()
            return
        
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                for line_num, line in enumerate(f, 1):
                    line = line.strip()
                    
                    # Skip empty lines and comments
                    if not line or line.startswith('#'):
                        continue
                    
                    # Parse key=value pairs
                    if '=' in line:
                        key, value = line.split('=', 1)
                        key = key.strip()
                        value = value.strip()
                        
                        # Convert values to appropriate types
                        self.config[key] = self._convert_value(value)
                    else:
                        print(f"Warning: Invalid line {line_num} in config file: {line}")
            
            print(f"Loaded configuration from {config_path}")
            
        except Exception as e:
            print(f"Error reading config file: {e}")
            self._set_defaults()
    
    def _convert_value(self, value: str) -> Any:
        """Convert string value to appropriate type."""
        # Boolean values
        if value.lower() in ('true', 'false'):
            return value.lower() == 'true'
        
        # Integer values
        try:
            return int(value)
        except ValueError:
            pass
        
        # Float values
        try:
            return float(value)
        except ValueError:
            pass
        
        # List values (comma-separated)
        if ',' in value:
            return [item.strip() for item in value.split(',') if item.strip()]
        
        # String values
        return value
    
    def _set_defaults(self):
        """Set default configuration values."""
        self.config = {
            'shared_mailbox_email': '',
            'shared_mailbox_name': 'Shared Mailbox',
            'personal_retention_months': 6,
            'shared_retention_months': 12,
            'max_search_results': 500,
            'max_body_chars': 0,
            'include_sent_items': True,
            'include_deleted_items': False,
            'connection_timeout_minutes': 10,
            'max_retry_attempts': 3,
            'batch_processing_size': 50,
            'analyze_importance_levels': True,
            'search_all_folders': False,
            'use_folder_traversal': False,
            'use_extended_mapi_login': True,
            'include_timestamps': True,
            'clean_html_content': True
        }
    
    def get(self, key: str, default=None):
        """Get configuration value by key."""
        return self.config.get(key, default)
    
    def get_int(self, key: str, default: int = 0) -> int:
        """Get configuration value as integer."""
        value = self.config.get(key, default)
        try:
            return int(value)
        except (ValueError, TypeError):
            return default
    
    def get_bool(self, key: str, default: bool = False) -> bool:
        """Get configuration value as boolean."""
        value = self.config.get(key, default)
        if isinstance(value, bool):
            return value
        if isinstance(value, str):
            return value.lower() in ('true', '1', 'yes', 'on')
        return default
    
    def get_list(self, key: str, default: List = None) -> List:
        """Get configuration value as list."""
        if default is None:
            default = []
        
        value = self.config.get(key, default)
        if isinstance(value, list):
            return value
        if isinstance(value, str):
            return [item.strip() for item in value.split(',') if item.strip()]
        return default
    
    def show_config(self):
        """Display current configuration."""
        print("\nCurrent Configuration:")
        print("=" * 40)
        for key, value in sorted(self.config.items()):
            # Don't show empty email addresses
            if key == 'shared_mailbox_email' and not value:
                print(f"{key}: <not configured>")
            else:
                print(f"{key}: {value}")
        print("=" * 40)


# Global config instance
config = ConfigReader()
