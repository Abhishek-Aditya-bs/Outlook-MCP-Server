# Outlook MCP Server

A Model Context Protocol (MCP) server that provides programmatic access to Microsoft Outlook mailboxes. This server enables AI assistants and MCP clients to search, analyze, and extract insights from emails in both personal and shared mailboxes.

## Features

- **Multi-Mailbox Support**: Access both personal inbox and shared mailboxes simultaneously
- **Advanced Search Capabilities**: Search emails by exact phrase matching in both subject and body
- **Near-Instant Search Performance**: Leverages Outlook's AdvancedSearch API for blazing-fast body searches
- **Parallel Mailbox Search**: Searches personal and shared mailboxes concurrently for faster results
- **Full Email Content**: Retrieves complete email bodies for comprehensive analysis
- **Email Chain Analysis**: Groups and analyzes related email conversations
- **Smart Connection Management**: Automatically connects to existing Outlook instances with retry logic
- **Optimized Caching**: Time-based cache invalidation with size limits for optimal memory usage
- **Non-Blocking Operations**: Async execution prevents server blocking during long operations
- **Configurable Settings**: Fine-tune performance and behavior through configuration file
- **Cross-Folder Search**: Optionally search across all folders, not just Inbox
- **Automatic Fallback**: Gracefully handles indexing issues with alternative search methods

## Requirements

- **Operating System**: Windows 10 or Windows 11
- **Microsoft Outlook**: Desktop application (not Outlook Web)
- **Python**: Version 3.8 or higher
- **Python Packages**:
  - `pywin32` - For Outlook COM interface
  - `mcp` - MCP SDK for server implementation

## Installation

1. **Clone the repository**:
```bash
git clone https://github.com/yourusername/outlook-mcp-server.git
cd outlook-mcp-server
```

2. **Install dependencies**:
```bash
pip install -r requirements.txt
```

3. **Configure settings**:
   - Edit `src/config/config.properties` with your mailbox details:
```properties
# Update this with your shared mailbox email (optional)
shared_mailbox_email=team-inbox@yourcompany.com

# Adjust search and performance settings as needed
max_search_results=50
max_body_chars=0  # 0 for full body, or limit for truncation
```

4. **Test the connection**:
```bash
python tests/test_connection.py
```

## Configuration

The server behavior can be customized through `config.properties`:

### Mailbox Settings
- `shared_mailbox_email`: Email address of shared/team mailbox (optional)
- `shared_mailbox_name`: Display name for the shared mailbox

### Search Configuration
- `max_search_results`: Maximum emails to return per search (default: 50)
- `max_body_chars`: Maximum characters from email body (0 = unlimited)
- `search_all_folders`: Search all folders, not just Inbox (default: false)

### Performance Settings
- `max_search_body_chars`: Limit for body searching during pattern matching
- `connection_timeout_minutes`: Outlook connection timeout
- `batch_processing_size`: Number of emails to process in batch
- `max_connection_retries`: Number of connection retry attempts (default: 3)
- `max_recipients_display`: Maximum recipients to show per email (default: 10)

### Data Retention (Informational)
- `personal_retention_months`: Expected retention for personal mailbox
- `shared_retention_months`: Expected retention for shared mailbox

## Usage

### Starting the Server

1. **Ensure Microsoft Outlook is running** on your system
2. **Start the MCP server**:
```bash
python outlook_mcp.py
```

3. The server will start and listen for MCP client connections via stdio

### Available Tools

The server provides two main tools accessible through the MCP protocol:

#### 1. `check_mailbox_access`
Tests connection to Outlook and verifies access to configured mailboxes.

**Parameters**: None

**Returns**: 
- Connection status
- Personal mailbox accessibility and name
- Shared mailbox accessibility and name (if configured)
- Any error messages

**Example Response**:
```json
{
  "status": "success",
  "connection": {
    "outlook_connected": true,
    "timestamp": "2024-01-15T10:30:00"
  },
  "personal_mailbox": {
    "accessible": true,
    "name": "John Doe",
    "retention_months": 6
  },
  "shared_mailbox": {
    "accessible": true,
    "name": "Team Support",
    "configured": true,
    "retention_months": 12
  }
}
```

#### 2. `get_email_chain`
Searches for emails containing specified text in both subject and body, returning complete email chains with full content.

**Parameters**:
- `search_text` (required): Exact phrase to search for
- `include_personal` (optional): Search personal mailbox (default: true)
- `include_shared` (optional): Search shared mailbox (default: true)

**Returns**:
- Grouped email conversations
- Full email bodies for each message
- Sender and recipient information
- Timestamps and folder locations
- Summary statistics

**Example Request**:
```json
{
  "tool": "get_email_chain",
  "arguments": {
    "search_text": "server error 500",
    "include_personal": true,
    "include_shared": true
  }
}
```

## Search Strategy

The server uses Outlook's AdvancedSearch API for near-instant search performance:

### Primary Search Method: AdvancedSearch API
- **Leverages Outlook's built-in search index** for blazing-fast performance
- **Searches both subject and body simultaneously** using DASL queries
- **Case-insensitive exact phrase matching** using `ci_phrasematch`
- **Near-instant results** even for large mailboxes (thousands of emails)
- **Asynchronous search** with polling for completion (30-second timeout)
- **Works identically to Outlook's UI search**, providing familiar behavior

### Automatic Fallback (if indexing is disabled)
If AdvancedSearch fails (rare, usually due to indexing issues):
1. **Subject-only search** using Restrict filters (always fast)
2. **Manual iteration** as last resort (limited scope)

### Other Folders Search (Optional)
- Searches Sent Items and Drafts using same AdvancedSearch method
- Activated when `search_all_folders=true`
- Consistent performance across all folders

## Performance Considerations

### Search Performance

**AdvancedSearch Benefits**:
- **10-100x faster** than traditional iteration methods
- **Sub-second to few seconds** response time for most searches
- **Consistent performance** regardless of mailbox size
- **Same speed for body searches as subject searches**

**Parallel Search**:
- Personal and shared mailboxes are searched simultaneously using threading
- ~2x faster when searching multiple mailboxes
- Proper COM initialization per thread ensures stability

**Smart Connection**:
- Connects to existing Outlook instance first (GetActiveObject) - instant connection
- Falls back to launching new instance only if needed
- Exponential backoff retry (1s, 2s, 4s) for resilient connection

**Optimized Caching**:
- 1-hour cache lifetime with automatic invalidation
- Cache size limited to 100 entries with LRU eviction
- Cache key includes search parameters for accuracy

**Memory Management**:
- COM references released after email extraction
- Recipients list limited to 10 by default (configurable)
- Email body truncation supported via max_body_chars

**Non-Blocking Server**:
- Uses asyncio.to_thread() for all Outlook operations
- Server remains responsive during long searches
- Multiple concurrent tool calls supported

**`max_results` Behavior**: The `max_results` configuration sets the total maximum number of emails returned across ALL mailboxes. Results are limited early during search for efficiency.

### Optimization Tips

1. **Ensure Outlook Indexing is Enabled**: 
   - Go to File → Options → Search → Indexing Options
   - Make sure Outlook is included in indexed locations
   - Allow indexing to complete for best performance

2. **Use Specific Search Terms**: More specific phrases yield faster, more accurate results

3. **Limit Results**: Set reasonable `max_search_results` to improve response times

4. **Configure Body Limits**: Use `max_body_chars` if full email bodies aren't needed for initial processing

5. **Keep Outlook Updated**: Newer versions have better search indexing and performance

### Caching

The server includes built-in caching for:
- Search results (keyed by search term and mailbox selection)
- Folder references (to avoid repeated lookups)

Cache is maintained per server session and cleared on restart.

## Integration with MCP Clients

This server is compatible with any MCP client that supports the stdio transport. Common integrations include:

### Claude Desktop App
Add to your Claude configuration:
```json
{
  "mcpServers": {
    "outlook": {
      "command": "python",
      "args": ["path/to/outlook_mcp.py"]
    }
  }
}
```

### Custom MCP Clients
Connect to the server using the MCP SDK:
```python
from mcp import Client

client = Client()
client.connect_stdio(["python", "outlook_mcp.py"])
```

## Troubleshooting

### Common Issues

**"Outlook.Application" Error**
- Ensure Microsoft Outlook desktop is installed (not just web access)
- Outlook must be running before starting the server

**Security Dialog Appears**
- This is normal on first access
- Click "Allow" to grant the server access to Outlook
- Consider enabling `use_extended_mapi_login` in config

**Shared Mailbox Not Accessible**
- Verify you have permissions to the shared mailbox
- Check the email address is correct in config.properties
- Ensure the mailbox is added to your Outlook profile

**Search Returns No Results**
- Verify emails exist matching your search criteria
- Try broader search terms
- Increase `max_search_results` if needed
- Check that Outlook has finished indexing your mailbox

**Slow Search Performance**
- Use more specific search terms
- Reduce `max_search_results`
- Consider limiting search to specific mailboxes
- Ensure Outlook is not syncing or downloading messages

### Debug Mode

Enable detailed logging by setting the logging level:
```python
import logging
logging.basicConfig(level=logging.DEBUG)
```

## Project Structure

```
outlook-mcp-server/
├── outlook_mcp.py           # Main MCP server
├── requirements.txt          # Python dependencies
├── src/
│   ├── config/
│   │   ├── config_reader.py # Configuration management
│   │   └── config.properties # User settings
│   └── utils/
│       ├── outlook_client.py # Outlook COM interface
│       └── email_formatter.py # Response formatting
└── tests/
    └── test_connection.py    # Connection test utility
```

## Architecture

### Components

1. **MCP Server Framework** (`outlook_mcp.py`)
   - Implements MCP protocol specification
   - Handles tool registration and execution
   - Manages stdio communication

2. **Outlook Client** (`src/utils/outlook_client.py`)
   - COM interface to Microsoft Outlook
   - Implements search strategies
   - Manages mailbox connections
   - Handles caching

3. **Email Formatter** (`src/utils/email_formatter.py`)
   - Formats email data for AI consumption
   - Groups emails into conversations
   - Generates summaries and statistics

4. **Configuration Reader** (`src/config/config_reader.py`)
   - Loads and validates configuration
   - Provides type-safe config access
   - Supports environment variable overrides

### Data Flow

1. MCP client sends tool request → MCP server
2. Server validates request parameters
3. Outlook client executes search strategy
4. Email data is extracted and formatted
5. Response is serialized and returned to client

## Security Considerations

- **Local Access Only**: Server runs locally and accesses Outlook via COM
- **Permission Prompts**: Windows may show security dialogs for Outlook access
- **No Credentials Stored**: Uses current Windows user's Outlook profile
- **Configurable Scope**: Limit access to specific mailboxes via configuration

## Contributing

Contributions are welcome! Please:

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests if applicable
5. Submit a pull request

## License

MIT License - See LICENSE file for details

## Support

For issues, questions, or suggestions:
- Open an issue on GitHub
- Check existing issues for solutions
- Provide detailed error messages and configuration (without sensitive data)

## Acknowledgments

- Built on the [Model Context Protocol](https://modelcontextprotocol.io) specification
- Uses [pywin32](https://github.com/mhammond/pywin32) for Windows COM interface
- Inspired by the need for AI-powered email analysis
