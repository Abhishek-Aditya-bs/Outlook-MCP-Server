"""Simplified Outlook MCP Server with three main tools."""

import asyncio
import logging
import platform
import sys
from typing import Any, Sequence

# Check if running on Windows
if platform.system() != 'Windows':
    print("❌ Error: Outlook MCP Server requires Windows with Microsoft Outlook installed")
    print(f"   Current platform: {platform.system()}")
    print("\n📋 To use this server:")
    print("   1. Run on a Windows machine with Outlook installed")
    print("   2. Or use a Windows virtual machine")
    print("   3. Or access a remote Windows desktop")
    sys.exit(1)

from mcp import server, types
from mcp.server import Server
from mcp.server.stdio import stdio_server

try:
    from src.config.config_reader import config
    from src.utils.outlook_client import outlook_client
    from src.utils.email_formatter import format_mailbox_status, format_email_chain, format_alert_analysis
except ImportError as e:
    print(f"❌ Import Error: {e}")
    print("\n📋 Please install required dependencies:")
    print("   pip install -r requirements.txt")
    print("\nNote: pywin32 is required and only works on Windows")
    sys.exit(1)

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Create MCP server
app = Server("outlook-mcp-server")


@app.list_tools()
async def list_tools() -> list[types.Tool]:
    """List available MCP tools."""
    return [
        types.Tool(
            name="check_mailbox_access",
            description="Check connection status and access to personal and shared mailboxes with retention policy info",
            inputSchema={
                "type": "object",
                "properties": {},
                "required": []
            }
        ),
        types.Tool(
            name="get_email_chain",
            description="Get email chain by subject pattern. Searches ALL folders in both personal and shared mailboxes (entire available mailbox like Outlook native search)",
            inputSchema={
                "type": "object",
                "properties": {
                    "subject": {
                        "type": "string",
                        "description": "Subject pattern to search for"
                    },
                    "include_personal": {
                        "type": "boolean",
                        "description": "Search personal mailbox (default: true)",
                        "default": True
                    },
                    "include_shared": {
                        "type": "boolean", 
                        "description": "Search shared mailbox (default: true)",
                        "default": True
                    }
                },
                "required": ["subject"]
            }
        ),
        types.Tool(
            name="analyze_alerts",
            description="Analyze production alerts across mailboxes. Searches entire available mailbox like Outlook native search. Finds alert patterns and provides recommendations",
            inputSchema={
                "type": "object",
                "properties": {
                    "alert_pattern": {
                        "type": "string",
                        "description": "Alert pattern or identifier to search for"
                    },
                    "include_personal": {
                        "type": "boolean",
                        "description": "Search personal mailbox (default: true)",
                        "default": True
                    },
                    "include_shared": {
                        "type": "boolean",
                        "description": "Search shared mailbox (default: true)", 
                        "default": True
                    }
                },
                "required": ["alert_pattern"]
            }
        )
    ]


@app.call_tool()
async def call_tool(name: str, arguments: dict[str, Any]) -> Sequence[types.TextContent]:
    """Handle tool calls."""
    
    logger.info(f"Executing tool: {name}")
    
    try:
        if name == "check_mailbox_access":
            return await handle_check_mailbox_access()
            
        elif name == "get_email_chain":
            subject = arguments.get("subject")
            if not subject:
                raise ValueError("Subject parameter is required")
            
            include_personal = arguments.get("include_personal", True)
            include_shared = arguments.get("include_shared", True)
            
            return await handle_get_email_chain(subject, include_personal, include_shared)
            
        elif name == "analyze_alerts":
            alert_pattern = arguments.get("alert_pattern")
            if not alert_pattern:
                raise ValueError("Alert pattern parameter is required")
            
            include_personal = arguments.get("include_personal", True)
            include_shared = arguments.get("include_shared", True)
            
            return await handle_analyze_alerts(alert_pattern, include_personal, include_shared)
            
        else:
            raise ValueError(f"Unknown tool: {name}")
            
    except Exception as e:
        logger.error(f"Error in tool {name}: {e}")
        error_response = {
            "status": "error",
            "tool": name,
            "error": str(e),
            "message": f"Failed to execute {name}: {str(e)}"
        }
        return [types.TextContent(type="text", text=str(error_response))]


async def handle_check_mailbox_access():
    """Handle mailbox access check."""
    logger.info("Checking mailbox access...")
    
    try:
        # Check access to mailboxes
        access_result = outlook_client.check_access()
        
        # Format response
        formatted_result = format_mailbox_status(access_result)
        
        logger.info("Mailbox access check completed")
        return [types.TextContent(type="text", text=str(formatted_result))]
        
    except Exception as e:
        logger.error(f"Error checking mailbox access: {e}")
        error_response = {
            "status": "error",
            "message": f"Could not check mailbox access: {str(e)}",
            "troubleshooting": [
                "Make sure Outlook is running",
                "Grant permission when security dialog appears", 
                "Check network connectivity"
            ]
        }
        return [types.TextContent(type="text", text=str(error_response))]


async def handle_get_email_chain(subject: str, include_personal: bool, include_shared: bool):
    """Handle email chain retrieval."""
    logger.info(f"Searching for email chain: {subject}")
    
    try:
        # Search for emails
        emails = outlook_client.search_emails_by_subject(
            subject=subject,
            include_personal=include_personal, 
            include_shared=include_shared
        )
        
        # Format response
        formatted_result = format_email_chain(emails, subject)
        
        logger.info(f"Found {len(emails)} emails for subject: {subject}")
        return [types.TextContent(type="text", text=str(formatted_result))]
        
    except Exception as e:
        logger.error(f"Error retrieving email chain: {e}")
        error_response = {
            "status": "error", 
            "search_subject": subject,
            "message": f"Could not retrieve email chain: {str(e)}",
            "troubleshooting": [
                "Verify Outlook connection", 
                "Check subject pattern spelling",
                "Ensure mailboxes are accessible"
            ]
        }
        return [types.TextContent(type="text", text=str(error_response))]


async def handle_analyze_alerts(alert_pattern: str, include_personal: bool, include_shared: bool):
    """Handle alert analysis."""
    logger.info(f"Analyzing alerts for pattern: {alert_pattern}")
    
    try:
        # Search for alerts
        alerts = outlook_client.search_alerts(
            alert_pattern=alert_pattern,
            include_personal=include_personal,
            include_shared=include_shared
        )
        
        # Format response  
        formatted_result = format_alert_analysis(alerts, alert_pattern)
        
        logger.info(f"Found {len(alerts)} alerts for pattern: {alert_pattern}")
        return [types.TextContent(type="text", text=str(formatted_result))]
        
    except Exception as e:
        logger.error(f"Error analyzing alerts: {e}")
        error_response = {
            "status": "error",
            "search_pattern": alert_pattern, 
            "message": f"Could not analyze alerts: {str(e)}",
            "troubleshooting": [
                "Verify Outlook connection",
                "Check alert pattern spelling", 
                "Ensure shared mailbox is configured"
            ]
        }
        return [types.TextContent(type="text", text=str(error_response))]


@app.list_resources()
async def list_resources() -> list[types.Resource]:
    """List available resources."""
    return [
        types.Resource(
            uri="outlook-mcp://config",
            name="Current Configuration", 
            description="Show current configuration settings",
            mimeType="text/plain"
        )
    ]


@app.read_resource()
async def read_resource(uri: str) -> str:
    """Read resource content."""
    if uri == "outlook-mcp://config":
        config.show_config()
        return "Configuration displayed in console"
    else:
        raise ValueError(f"Unknown resource: {uri}")


async def main():
    """Main entry point."""
    print("=" * 60)
    print("🚀 Starting Outlook MCP Server")
    print("=" * 60)
    
    # Show configuration
    config.show_config()
    
    # Important notes
    print("\n📋 Important Notes:")
    print("   • Make sure Microsoft Outlook is running")
    print("   • Grant permission when security dialog appears")  
    print("   • Update config.properties with your shared mailbox details")
    print("   • Server searches ALL folders, not just Inbox")
    
    shared_email = config.get('shared_mailbox_email')
    if not shared_email or shared_email == 'your-shared-mailbox@company.com':
        print("\n⚠️  Warning: Shared mailbox not configured!")
        print("   Update 'shared_mailbox_email' in config.properties")
    
    print("\n🔧 Available Tools:")
    print("   1. check_mailbox_access - Test connection and access")
    print("   2. get_email_chain - Find email conversations by subject")
    print("   3. analyze_alerts - Analyze production alerts and patterns")
    
    print(f"\n✅ Server ready! Listening for MCP client connections...")
    print("=" * 60)
    
    # Start server
    async with stdio_server() as (read_stream, write_stream):
        await app.run(read_stream, write_stream, app.create_initialization_options())


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\n👋 Server stopped by user")
    except Exception as e:
        print(f"\n❌ Server error: {e}")
        logger.error(f"Server error: {e}")
