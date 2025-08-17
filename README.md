# Outlook MCP Server

A simplified Model Context Protocol (MCP) server for Microsoft Outlook integration. Provides AI assistants with access to email analysis, alert monitoring, and conversation tracking.

## Features

**Simple Setup** - Just 5 files, easy configuration  
**Three Tools** - Access, email chains, and alert analysis  
**All Folders** - Searches entire mailbox, not just Inbox  
**Dual Mailbox** - Personal and shared mailbox support  
**AI Ready** - Structured responses for AI analysis

## Quick Start

### 1. Requirements
- Windows 10 or Windows 11
- Microsoft Outlook installed and running
- Python 3.8+

### 2. Install
```bash
git clone <your-repo-url>
cd Outlook-MCP-Server
pip install -r requirements.txt
```

### 3. Configure
Edit `src/config/config.properties`:
```properties
# Update with your details
shared_mailbox_email=production-alerts@yourcompany.com
shared_mailbox_name=Production Monitoring
personal_retention_months=6
shared_retention_months=12
```

### 4. Test Connection
```bash
python tests/test_connection.py
```

### 5. Start Server
```bash
python outlook_mcp.py
```

## Project Structure

```
Outlook-MCP-Server/
├── src/
│   ├── config/
│   │   ├── config.properties      # Your configuration (Java-style)
│   │   └── config_reader.py       # Properties file reader
│   └── utils/
│       ├── outlook_client.py      # Outlook COM interface
│       └── email_formatter.py     # AI-optimized formatting
├── tests/
│   ├── test_connection.py         # Simple test script
│   └── test-shared-mailbox.py     # Original test file
├── docs/
│   └── PROJECT_OVERVIEW.md        # Detailed project documentation
├── outlook_mcp.py                 # Main MCP server (3 tools)
├── requirements.txt               # Dependencies
└── README.md                      # This file
```

## Three Tools

### 1. Check Mailbox Access
```python
# Verifies connection and shows retention policies
check_mailbox_access()
```

**Returns:**
- Connection status to Outlook
- Personal mailbox access (with retention policy info)
- Shared mailbox access (with retention policy info)
- Error diagnostics and recommendations

**Note**: Retention policies are informational only - searches cover the entire available mailbox regardless of policy

### 2. Get Email Chain
```python
# Find email conversations by subject (searches entire available mailbox)
get_email_chain(
    subject="Production Issue Database",
    include_personal=True,  # Optional
    include_shared=True     # Optional
)
```

**Returns:**
- Grouped email conversations
- Participant analysis  
- Timeline and chronology
- Cross-mailbox tracking

### 3. Analyze Alerts
```python
# Analyze production alerts and patterns (searches entire available mailbox)
analyze_alerts(
    alert_pattern="database timeout", 
    include_personal=True,  # Optional  
    include_shared=True     # Optional
)
```

**Returns:**
- Alert frequency analysis
- Urgent vs normal alerts
- Response pattern analysis
- Actionable recommendations

## Configuration Options

Edit `src/config/config.properties` for your organization:

```properties
# === Mailbox Settings ===
shared_mailbox_email=alerts@company.com
shared_mailbox_name=Production Alerts
personal_retention_months=6          # Informational only
shared_retention_months=12           # Informational only

# === Search Settings ===
max_search_results=500
max_body_chars=0                    # 0 = no limit (full email body)
search_all_folders=true
include_sent_items=true
include_deleted_items=false

# === Alert Analysis ===
analyze_importance_levels=true      # Use email importance for urgency detection

# === Security Settings ===
use_extended_mapi_login=true         # Try Extended MAPI to reduce security prompts
```

**Note on `max_body_chars`**: 
- `0` = Include full email body (recommended for AI analysis)
- `5000` = Limit to 5000 characters (for performance with very large emails)
- Large emails may impact processing speed but provide complete context

## Usage with MCP Clients

### GitHub Copilot / VS Code
Add to your MCP configuration:
```json
{
  "mcpServers": {
    "outlook": {
      "command": "python",
      "args": ["outlook_mcp.py"], 
      "cwd": "/path/to/Outlook-MCP-Server"
    }
  }
}
```

### Claude Desktop
```json
{
  "mcpServers": {
    "outlook": {
      "command": "python",
      "args": ["/path/to/Outlook-MCP-Server/outlook_mcp.py"]
    }
  }
}
```

## Security Notes

**Outlook Permission Dialog**: When first connecting, you'll see:
```
"A program is trying to access e-mail addresses..."
```

**Action Required**: Click "Allow access for 10 minutes" (or desired duration)

**Why This Happens**: Windows security feature for COM access to Outlook

**Frequency**: You may need to grant permission periodically

**Security Bypass Attempt**: The server tries an Extended MAPI login approach that may reduce (but not eliminate) security prompts. This is enabled by default in `src/config/config.properties`

## Troubleshooting

### "Failed to connect to Outlook"
- Start Microsoft Outlook
- Log in to your account  
- Grant permission when dialog appears

### "Still getting security prompts frequently"
- Ensure `use_extended_mapi_login=true` in src/config/config.properties
- Try granting longer access duration (e.g., 10 minutes instead of 1 minute)
- Note: Complete elimination of prompts is not guaranteed due to Outlook security

### "Shared mailbox not accessible"
- Update `shared_mailbox_email` in src/config/config.properties
- Ensure you have delegate access rights
- Add shared mailbox to your Outlook profile

### "No emails found"
- Check search terms are correct
- Verify emails exist in date range
- Try broader search patterns
- Check retention policy settings

### Performance Issues  
- Lower `max_search_results` in config
- Set `search_all_folders=false` to search only Inbox
- Set `max_body_chars=2000` to limit email body size

## Development

### Run Tests
```bash
python tests/test_connection.py
```

### Debug Mode
```bash
# Add to outlook_mcp.py
import logging
logging.basicConfig(level=logging.DEBUG)
```

### Customize Settings
```properties
# Edit src/config/config.properties
max_body_chars=5000              # Limit email body size
analyze_importance_levels=false  # Disable urgency detection
search_all_folders=false         # Search only Inbox
```

## License

MIT License - See LICENSE file

## Support

1. Check troubleshooting section above
2. Run `python tests/test_connection.py` for diagnostics  
3. Verify `src/config/config.properties` settings
4. Ensure Outlook is running and accessible

---

**Note**: This server only works on Windows with Microsoft Outlook installed. It uses the COM interface to access mailbox data locally - no data is transmitted externally.