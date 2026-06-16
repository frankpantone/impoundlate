#!/usr/bin/env python3
"""
Simple script to send a test email using Outlook COM interface.
Make sure Outlook is running before executing this script.
"""

from outlook_email_sender import OutlookEmailSender
import sys

def send_test_email():
    """Send a test email to verify the setup."""
    try:
        # Create email sender instance
        print("🔗 Connecting to Outlook...")
        email_sender = OutlookEmailSender()
        
        # Get current user's email
        current_email = email_sender.get_current_user_email()
        if current_email:
            print(f"✅ Connected to Outlook account: {current_email}")
        else:
            print("⚠️  Could not retrieve current user email address")
        
        # Ask user for test email address
        print("\n📧 Test Email Setup")
        print("=" * 30)
        
        # Option 1: Send to yourself
        print("1. Send test email to yourself")
        print("2. Send test email to another address")
        
        choice = input("\nEnter your choice (1 or 2): ").strip()
        
        if choice == "1":
            if current_email:
                recipient = current_email
                print(f"📤 Sending test email to: {recipient}")
            else:
                print("❌ Cannot send to yourself - email address not detected")
                recipient = input("Please enter your email address: ").strip()
        elif choice == "2":
            recipient = input("Enter recipient email address: ").strip()
        else:
            print("❌ Invalid choice")
            return
        
        if not recipient:
            print("❌ No recipient email address provided")
            return
        
        # Send test email
        print(f"\n📬 Sending test email to {recipient}...")
        
        if email_sender.send_test_email(recipient):
            print("✅ Test email sent successfully!")
            print("📬 Check your inbox to confirm delivery.")
        else:
            print("❌ Failed to send test email")
            print("💡 Make sure Outlook is running and you have an active internet connection")
            
    except Exception as e:
        print(f"❌ Error: {e}")
        print("\n💡 Troubleshooting tips:")
        print("   • Make sure Microsoft Outlook is running")
        print("   • Ensure you have an active email account configured in Outlook")
        print("   • Check that pywin32 is installed: pip install pywin32")
        sys.exit(1)

def send_custom_email():
    """Send a custom email with user-provided content."""
    try:
        email_sender = OutlookEmailSender()
        
        print("\n📝 Custom Email Setup")
        print("=" * 25)
        
        # Get email details from user
        recipients = []
        print("Enter recipient email addresses (press Enter when done):")
        while True:
            email = input("Recipient: ").strip()
            if not email:
                break
            recipients.append(email)
        
        if not recipients:
            print("❌ No recipients provided")
            return
        
        subject = input("Subject: ").strip()
        if not subject:
            subject = "Email from Python Automation"
        
        print("Email body (press Ctrl+Z then Enter on Windows, or Ctrl+D on Unix when done):")
        body_lines = []
        try:
            while True:
                line = input()
                body_lines.append(line)
        except EOFError:
            pass
        
        body = "\n".join(body_lines)
        if not body:
            body = "This email was sent using Python automation."
        
        # Send the email
        print(f"\n📤 Sending email to {len(recipients)} recipient(s)...")
        
        if email_sender.send_email(recipients, subject, body):
            print("✅ Email sent successfully!")
        else:
            print("❌ Failed to send email")
            
    except Exception as e:
        print(f"❌ Error: {e}")

def main():
    """Main function to choose between test and custom email."""
    print("🚀 Outlook Email Automation")
    print("=" * 30)
    print("1. Send test email")
    print("2. Send custom email")
    print("3. Exit")
    
    choice = input("\nEnter your choice (1-3): ").strip()
    
    if choice == "1":
        send_test_email()
    elif choice == "2":
        send_custom_email()
    elif choice == "3":
        print("👋 Goodbye!")
        return
    else:
        print("❌ Invalid choice")

if __name__ == "__main__":
    main()







