#!/usr/bin/env python3
"""
Check for Shared Mailboxes and Send-As Permissions in Outlook
"""

import win32com.client
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


def check_shared_mailboxes():
    """Check for shared mailboxes and send-as permissions."""
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        print("="*80)
        print("CHECKING FOR SHARED MAILBOXES AND SEND-AS PERMISSIONS")
        print("="*80)
        
        # Check primary accounts
        accounts = namespace.Accounts
        print(f"\n[INFO] Primary Accounts: {accounts.Count}")
        for i in range(1, accounts.Count + 1):
            account = accounts.Item(i)
            print(f"  - {account.SmtpAddress}")
        
        # Check all folders (including shared mailboxes)
        print("\n[INFO] Checking all available folders and mailboxes...")
        
        folders = namespace.Folders
        print(f"\n[INFO] Found {folders.Count} folder(s)/mailbox(es):")
        
        shared_mailboxes = []
        for i in range(1, folders.Count + 1):
            folder = folders.Item(i)
            folder_name = folder.Name
            print(f"\n{i}. Folder/Mailbox: {folder_name}")
            
            # Try to get store properties
            try:
                store = folder.Store
                if store:
                    print(f"   Store Display Name: {store.DisplayName}")
                    
                    # Check if it's an Exchange store with additional mailbox
                    try:
                        # Try to access the Inbox to verify it's a mailbox
                        inbox = folder.Folders("Inbox")
                        if inbox:
                            print(f"   [INFO] This appears to be a mailbox (has Inbox)")
                            
                            # Check if it's different from primary account
                            if folder_name.lower() not in [acc.SmtpAddress.lower().split('@')[0] for acc in [accounts.Item(j) for j in range(1, accounts.Count + 1)]]:
                                shared_mailboxes.append(folder_name)
                                print(f"   [SUCCESS] Possible shared mailbox detected!")
                    except:
                        pass
            except Exception as e:
                logger.debug(f"Could not get store info for {folder_name}: {e}")
        
        # Check specifically for dispatch@hertz.com
        print("\n" + "="*80)
        print("SEARCHING FOR 'dispatch@hertz.com'")
        print("="*80)
        
        found_dispatch = False
        
        # Check in accounts
        for i in range(1, accounts.Count + 1):
            account = accounts.Item(i)
            if 'dispatch@hertz.com' in account.SmtpAddress.lower():
                print(f"\n[SUCCESS] Found dispatch@hertz.com in accounts!")
                print(f"  Account: {account.SmtpAddress}")
                found_dispatch = True
        
        # Check in folder names
        for i in range(1, folders.Count + 1):
            folder = folders.Item(i)
            if 'dispatch' in folder.Name.lower():
                print(f"\n[INFO] Found folder with 'dispatch' in name:")
                print(f"  Folder: {folder.Name}")
                found_dispatch = True
        
        if not found_dispatch:
            print("\n[WARNING] 'dispatch@hertz.com' not found in Outlook")
            print("\n[INFO] To use dispatch@hertz.com, you need to:")
            print("  1. Open Outlook Desktop App")
            print("  2. Go to File > Account Settings > Account Settings")
            print("  3. Select your account > Change > More Settings > Advanced")
            print("  4. Click 'Add' and enter: dispatch@hertz.com")
            print("  5. Click OK and restart Outlook")
            print("\nOR if you have delegate/send-as permissions:")
            print("  - The mailbox should automatically appear in your folder list")
            print("  - You may need to contact IT to verify permissions")
        
        # Test creating email with custom From address
        print("\n" + "="*80)
        print("TESTING SEND-AS CAPABILITY")
        print("="*80)
        
        try:
            mail = outlook.CreateItem(0)
            mail.To = "test@example.com"
            mail.Subject = "Test - Not Sent"
            mail.Body = "Test"
            
            # Try to set SentOnBehalfOfName
            try:
                mail.SentOnBehalfOfName = "dispatch@hertz.com"
                print("\n[INFO] Can set 'Send on behalf of' to dispatch@hertz.com")
                print("  (This will show as 'matthew.biocchi@hertz.com on behalf of dispatch@hertz.com')")
            except Exception as e:
                print(f"\n[WARNING] Cannot set 'Send on behalf of': {e}")
            
            mail = None
            print("\n[INFO] Test complete (no email sent)")
            
        except Exception as e:
            print(f"[ERROR] Email test failed: {e}")
        
        print("\n" + "="*80)
        print("RECOMMENDATION")
        print("="*80)
        
        if found_dispatch:
            print("[SUCCESS] dispatch@hertz.com is accessible")
            print("[INFO] You should be able to send emails from dispatch@hertz.com")
        else:
            print("[WARNING] dispatch@hertz.com is not configured in Outlook")
            print("\nOptions:")
            print("  A) Add dispatch@hertz.com as a shared mailbox (recommended)")
            print("  B) Send from matthew.biocchi@hertz.com instead")
            print("  C) Send 'on behalf of' dispatch@hertz.com (shows both addresses)")
        
    except Exception as e:
        print(f"[ERROR] Failed to check: {e}")
        logger.error(f"Error: {e}", exc_info=True)


if __name__ == "__main__":
    check_shared_mailboxes()




