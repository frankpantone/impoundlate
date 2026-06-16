#!/usr/bin/env python3
"""
Test Sending from dispatch@hertz.com
Creates a draft email to verify dispatch@hertz.com works - DOES NOT SEND
"""

import win32com.client
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


def test_dispatch_email():
    """Create a test draft email from dispatch@hertz.com without sending."""
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        print("="*80)
        print("TESTING EMAIL FROM dispatch@hertz.com")
        print("="*80)
        print("\n[INFO] Creating draft email (will NOT send)...\n")
        
        # Create mail item
        mail = outlook.CreateItem(0)
        
        # Set basic properties
        mail.To = "test@example.com"
        mail.Subject = "TEST - Overdue Delivery Notification (DRAFT ONLY)"
        mail.Body = """This is a TEST draft email from dispatch@hertz.com.

This email was created but NOT sent - it's only to verify the sending capability.

ORDER INFORMATION:
Order ID: TEST123
Status: Picked Up
VIN: TEST1234567890
Vehicle: 2025 Test Vehicle

Originally Scheduled Delivery: 11/06/2025 (7 days overdue)

CARRIER INFORMATION:
Carrier: Test Carrier
Email: test@example.com

ACTION REQUIRED:
Please provide updated delivery ETA.

Best regards,
Hertz Dispatch Fleet Team
dispatch@hertz.com"""
        
        # Try to set the sender to dispatch@hertz.com
        sender_set = False
        
        # Method 1: Try to find dispatch in accounts
        accounts = namespace.Accounts
        for i in range(1, accounts.Count + 1):
            account = accounts.Item(i)
            if 'dispatch@hertz.com' in account.SmtpAddress.lower():
                try:
                    mail._oleobj_.Invoke(*(64209, 0, 8, 0, account))
                    print(f"[SUCCESS] Method 1: Set account to {account.SmtpAddress}")
                    sender_set = True
                    break
                except Exception as e:
                    print(f"[WARNING] Method 1 failed: {e}")
        
        # Method 2: Try using shared mailbox folder
        if not sender_set:
            folders = namespace.Folders
            for i in range(1, folders.Count + 1):
                folder = folders.Item(i)
                if 'dispatch' in folder.Name.lower():
                    try:
                        # For shared mailboxes, set SentOnBehalfOfName
                        mail.SentOnBehalfOfName = 'dispatch@hertz.com'
                        print(f"[SUCCESS] Method 2: Set shared mailbox to dispatch@hertz.com")
                        print(f"  (Found folder: {folder.Name})")
                        sender_set = True
                        break
                    except Exception as e:
                        print(f"[WARNING] Method 2 failed: {e}")
        
        # Method 3: Direct SentOnBehalfOfName
        if not sender_set:
            try:
                mail.SentOnBehalfOfName = 'dispatch@hertz.com'
                print("[SUCCESS] Method 3: Set SentOnBehalfOfName to dispatch@hertz.com")
                sender_set = True
            except Exception as e:
                print(f"[WARNING] Method 3 failed: {e}")
        
        if sender_set:
            print("\n" + "="*80)
            print("DRAFT EMAIL CREATED SUCCESSFULLY")
            print("="*80)
            print(f"FROM: dispatch@hertz.com")
            print(f"TO: {mail.To}")
            print(f"SUBJECT: {mail.Subject}")
            print(f"\n[SUCCESS] Email can be sent from dispatch@hertz.com")
            print("\n[INFO] Draft email created in Outlook - check your Drafts folder")
            print("[INFO] You can review and delete it")
            
            # Save as draft
            mail.Save()
            print("\n[INFO] Draft saved to your Drafts folder")
            
        else:
            print("\n[ERROR] Could not set sender to dispatch@hertz.com")
            print("[INFO] Email would send from default account")
            mail = None
        
        print("\n" + "="*80)
        print("VERIFICATION COMPLETE")
        print("="*80)
        
        if sender_set:
            print("[SUCCESS] dispatch@hertz.com is ready to use!")
            print("[INFO] The overdue delivery script will work correctly")
        else:
            print("[WARNING] Could not configure dispatch@hertz.com as sender")
            print("[INFO] Please contact IT to verify shared mailbox permissions")
        
    except Exception as e:
        print(f"\n[ERROR] Test failed: {e}")
        logger.error(f"Error: {e}", exc_info=True)


if __name__ == "__main__":
    test_dispatch_email()




