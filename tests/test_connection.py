"""Simple test script to verify Outlook MCP Server setup."""

import asyncio
import sys
import os

# Add src directory to path for imports  
sys.path.append(os.path.join(os.path.dirname(__file__), '..', 'src'))

from src.utils.outlook_client import outlook_client
from src.config.config_reader import config
from src.utils.email_formatter import format_mailbox_status, format_email_chain


async def test_connection():
    """Test Outlook connection and basic functionality."""
    
    print("🔧 Outlook MCP Server - Connection Test")
    print("=" * 50)
    
    # Show current configuration
    print("\n📋 Current Configuration:")
    config.show_config()
    
    print("\n1️⃣  Testing Outlook Connection...")
    print("-" * 30)
    
    try:
        # Test mailbox access
        access_result = outlook_client.check_access()
        formatted_result = format_mailbox_status(access_result)
        
        # Display results
        connection = formatted_result["connection"]
        personal = formatted_result["personal_mailbox"] 
        shared = formatted_result["shared_mailbox"]
        
        print(f"   Outlook Connected: {'✅' if connection['outlook_connected'] else '❌'}")
        print(f"   Personal Mailbox: {'✅' if personal['accessible'] else '❌'} ({personal.get('name', 'Unknown')})")
        print(f"   Shared Mailbox: {'✅' if shared['accessible'] else '❌'} ({shared.get('name', 'Not configured')})")
        
        if formatted_result.get("errors"):
            print(f"   ⚠️  Errors: {len(formatted_result['errors'])}")
            for error in formatted_result["errors"]:
                print(f"      • {error}")
        
        connection_ok = connection["outlook_connected"] and personal["accessible"]
        
    except Exception as e:
        print(f"   ❌ Connection test failed: {e}")
        print("   💡 Make sure Outlook is running and grant permission when prompted")
        connection_ok = False
    
    if not connection_ok:
        print("\n❌ Connection test failed. Please resolve issues before continuing.")
        return
    
    print("\n2️⃣  Testing Email Search...")
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
            print(f"   ✅ Search successful!")
            print(f"   📧 Found {summary['total_emails']} emails in {summary['conversations']} conversations")
            print(f"   📁 Mailbox distribution: {summary['mailbox_distribution']}")
            
            if summary["total_emails"] > 0:
                date_range = summary["date_range"]
                print(f"   📅 Date range: {date_range['first'][:10]} to {date_range['last'][:10]}")
        else:
            print(f"   ℹ️  No emails found for '{test_subject}'")
            print("   💡 Try a different search term or check if emails exist in your mailbox")
        
    except Exception as e:
        print(f"   ❌ Email search failed: {e}")
        return
    
    print("\n3️⃣  Testing Alert Analysis...")
    print("-" * 30)
    
    try:
        # Test alert analysis with a simple alert term
        test_pattern = "alert"
        
        alerts = outlook_client.search_alerts(
            alert_pattern=test_pattern,
            include_personal=True,
            include_shared=True
        )
        
        print(f"   ✅ Alert search completed!")
        print(f"   🚨 Found {len(alerts)} potential alerts for pattern '{test_pattern}'")
        
        if alerts:
            # Show recent alerts
            recent_alerts = sorted(alerts, key=lambda x: x.get('received_time', ''), reverse=True)[:3]
            print(f"   📋 Recent alerts:")
            for i, alert in enumerate(recent_alerts, 1):
                subject = alert.get('subject', 'No Subject')
                sender = alert.get('sender_name', 'Unknown')
                print(f"      {i}. {subject[:60]}... (from {sender})")
        
    except Exception as e:
        print(f"   ❌ Alert analysis failed: {e}")
        return
    
    print("\n" + "=" * 50)
    print("🎉 All tests completed successfully!")
    print("=" * 50)
    
    print("\n✅ Your Outlook MCP Server is ready to use!")
    print("\n🚀 Next steps:")
    print("   1. Start the MCP server: python outlook_mcp.py")
    print("   2. Configure your MCP client to connect to this server")  
    print("   3. Update config.properties with your organization's details")
    
    # Configuration reminders
    shared_email = config.get('shared_mailbox_email', '')
    if not shared_email or 'your-shared-mailbox' in shared_email:
        print("\n⚠️  Don't forget to:")
        print("   • Update shared_mailbox_email in config.properties")
        print("   • Set appropriate retention policies")


async def main():
    """Main test function."""
    try:
        await test_connection()
    except KeyboardInterrupt:
        print("\n\n👋 Test interrupted by user")
    except Exception as e:
        print(f"\n\n❌ Unexpected error: {e}")
        print("Please check your setup and try again")


if __name__ == "__main__":
    print("Make sure Microsoft Outlook is running before starting this test...")
    input("Press Enter to continue...")
    print()
    
    asyncio.run(main())
