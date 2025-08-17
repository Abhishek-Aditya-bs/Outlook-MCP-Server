"""Simple test script to verify Outlook MCP Server setup."""

import asyncio
import sys
import os
import platform

# Check if running on Windows
if platform.system() != 'Windows':
    print("[ERROR] Outlook MCP Server requires Windows with Microsoft Outlook installed")
    print(f"   Current platform: {platform.system()}")
    print("\n[INFO] To use this server:")
    print("   1. Run on a Windows machine with Outlook installed")
    print("   2. Or use a Windows virtual machine")
    print("   3. Or access a remote Windows desktop")
    sys.exit(1)

# Add parent directory to path for imports
parent_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, parent_path)

try:
    from src.utils.outlook_client import outlook_client
    from src.config.config_reader import config
    from src.utils.email_formatter import format_mailbox_status, format_email_chain
except ImportError as e:
    print(f"[ERROR] Import Error: {e}")
    print("\n[INFO] Please install required dependencies:")
    print("   pip install -r requirements.txt")
    print("\nNote: pywin32 is required and only works on Windows")
    sys.exit(1)


async def test_connection():
    """Test Outlook connection and basic functionality."""
    
    print("[TEST] Outlook MCP Server - Connection Test")
    print("=" * 50)
    
    # Show current configuration
    print("\n[CONFIG] Current Configuration:")
    config.show_config()
    
    print("\n[1] Testing Outlook Connection...")
    print("-" * 30)
    
    try:
        # Test mailbox access
        access_result = outlook_client.check_access()
        formatted_result = format_mailbox_status(access_result)
        
        # Display results
        connection = formatted_result["connection"]
        personal = formatted_result["personal_mailbox"] 
        shared = formatted_result["shared_mailbox"]
        
        print(f"   Outlook Connected: {'[OK]' if connection['outlook_connected'] else '[FAIL]'}")
        print(f"   Personal Mailbox: {'[OK]' if personal['accessible'] else '[FAIL]'} ({personal.get('name', 'Unknown')})")
        print(f"   Shared Mailbox: {'[OK]' if shared['accessible'] else '[FAIL]'} ({shared.get('name', 'Not configured')})")
        
        if formatted_result.get("errors"):
            print(f"   [WARNING] Errors: {len(formatted_result['errors'])}")
            for error in formatted_result["errors"]:
                print(f"      * {error}")
        
        connection_ok = connection["outlook_connected"] and personal["accessible"]
        
    except Exception as e:
        print(f"   [ERROR] Connection test failed: {e}")
        print("   [TIP] Make sure Outlook is running and grant permission when prompted")
        connection_ok = False
    
    if not connection_ok:
        print("\n[ERROR] Connection test failed. Please resolve issues before continuing.")
        return
    
    print("\n[2] Testing Email Search...")
    print("-" * 30)
    
    # Test with a simple search
    test_subject = input("   Enter a subject to search for (or press Enter for 'test'): ").strip()
    if not test_subject:
        test_subject = "test"
    
    try:
        emails = outlook_client.search_emails_by_subject(
            subject=test_subject,
            include_personal=True,
            include_shared=True
        )
        
        formatted_result = format_email_chain(emails, test_subject)
        
        if formatted_result["status"] == "success":
            summary = formatted_result["summary"]
            print(f"   [OK] Search successful!")
            print(f"   [EMAIL] Found {summary['total_emails']} emails in {summary['conversations']} conversations")
            print(f"   [FOLDER] Mailbox distribution: {summary['mailbox_distribution']}")
            
            if summary["total_emails"] > 0:
                date_range = summary["date_range"]
                print(f"   [DATE] Date range: {date_range['first'][:10]} to {date_range['last'][:10]}")
        else:
            print(f"   [INFO] No emails found for '{test_subject}'")
            print("   [TIP] Try a different search term or check if emails exist in your mailbox")
        
    except Exception as e:
        print(f"   [ERROR] Email search failed: {e}")
        return
    
    print("\n[3] Testing Alert Analysis...")
    print("-" * 30)
    
    try:
        # Test alert analysis with a simple alert term
        test_pattern = "alert"
        
        alerts = outlook_client.search_alerts(
            alert_pattern=test_pattern,
            include_personal=True,
            include_shared=True
        )
        
        print(f"   [OK] Alert search completed!")
        print(f"   [ALERT] Found {len(alerts)} potential alerts for pattern '{test_pattern}'")
        
        if alerts:
            # Show recent alerts
            recent_alerts = sorted(alerts, key=lambda x: x.get('received_time', ''), reverse=True)[:3]
            print(f"   [RECENT] Recent alerts:")
            for i, alert in enumerate(recent_alerts, 1):
                subject = alert.get('subject', 'No Subject')
                sender = alert.get('sender_name', 'Unknown')
                print(f"      {i}. {subject[:60]}... (from {sender})")
        
    except Exception as e:
        print(f"   [ERROR] Alert analysis failed: {e}")
        return
    
    print("\n" + "=" * 50)
    print("[SUCCESS] All tests completed successfully!")
    print("=" * 50)
    
    print("\n[READY] Your Outlook MCP Server is ready to use!")
    print("\n[NEXT STEPS]:")
    print("   1. Start the MCP server: python outlook_mcp.py")
    print("   2. Configure your MCP client to connect to this server")  
    print("   3. Update config.properties with your organization's details")
    
    # Configuration reminders
    shared_email = config.get('shared_mailbox_email', '')
    if not shared_email or 'your-shared-mailbox' in shared_email:
        print("\n[REMINDER] Don't forget to:")
        print("   * Update shared_mailbox_email in config.properties")
        print("   * Set appropriate retention policies")


async def main():
    """Main test function."""
    try:
        await test_connection()
    except KeyboardInterrupt:
        print("\n\n[INFO] Test interrupted by user")
    except Exception as e:
        print(f"\n\n[ERROR] Unexpected error: {e}")
        print("Please check your setup and try again")


if __name__ == "__main__":
    print("Make sure Microsoft Outlook is running before starting this test...")
    input("Press Enter to continue...")
    print()
    
    asyncio.run(main())
