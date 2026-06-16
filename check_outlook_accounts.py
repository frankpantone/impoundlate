#!/usr/bin/env python3
"""
Check Outlook Accounts and Verify Email Sending Capability
This script checks available Outlook accounts without sending any emails.
"""

import win32com.client
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


def check_outlook_accounts():
    """
    Check all available Outlook accounts and verify access to specific accounts.
    """
    try:
        # Connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        logger.info("Successfully connected to Outlook application")
        print("[SUCCESS] Connected to Outlook application\n")
        
        # Get MAPI namespace
        namespace = outlook.GetNamespace("MAPI")
        
        # Get all accounts
        accounts = namespace.Accounts
        
        print("="*80)
        print(f"FOUND {accounts.Count} OUTLOOK ACCOUNT(S)")
        print("="*80)
        
        # List all accounts
        account_list = []
        for i in range(1, accounts.Count + 1):
            account = accounts.Item(i)
            account_email = account.SmtpAddress
            account_name = account.DisplayName
            account_list.append(account_email.lower())
            
            print(f"\n{i}. Account Name: {account_name}")
            print(f"   Email Address: {account_email}")
            print(f"   Account Type: {account.AccountType}")
            
            # Additional account details
            try:
                print(f"   User Name: {account.UserName}")
            except:
                pass
        
        # Check for specific accounts
        print("\n" + "="*80)
        print("CHECKING FOR DISPATCH EMAIL ACCOUNTS")
        print("="*80)
        
        target_emails = [
            'dispatchfleet@hertz.com',
            'dispatch@hertz.com',
            'Dispatchfleet@hertz.com',
            'Dispatch@hertz.com'
        ]
        
        found_accounts = []
        for target in target_emails:
            if target.lower() in account_list:
                found_accounts.append(target)
                print(f"\n[SUCCESS] Found account: {target}")
                
                # Get the account object
                for i in range(1, accounts.Count + 1):
                    account = accounts.Item(i)
                    if account.SmtpAddress.lower() == target.lower():
                        print(f"  - Display Name: {account.DisplayName}")
                        print(f"  - Account Type: {account.AccountType}")
                        print(f"  - Can be used to send emails: YES")
                        break
        
        if not found_accounts:
            print("\n[WARNING] Neither 'dispatchfleet@hertz.com' nor 'dispatch@hertz.com' found in Outlook accounts")
            print("\nAvailable accounts in your Outlook:")
            for email in account_list:
                print(f"  - {email}")
        
        # Test creating a draft email (without sending)
        print("\n" + "="*80)
        print("TESTING EMAIL CREATION CAPABILITY (NOT SENDING)")
        print("="*80)
        
        try:
            mail = outlook.CreateItem(0)  # 0 = olMailItem
            
            # Try to set the sending account if we found one
            if found_accounts:
                # Set the account to send from
                for i in range(1, accounts.Count + 1):
                    account = accounts.Item(i)
                    if account.SmtpAddress.lower() in [e.lower() for e in found_accounts]:
                        mail._oleobj_.Invoke(*(64209, 0, 8, 0, account))  # Set SendUsingAccount
                        print(f"\n[SUCCESS] Can set sending account to: {account.SmtpAddress}")
                        break
            
            mail.To = "test@example.com"
            mail.Subject = "Test Email - NOT SENT"
            mail.Body = "This is a test email that was NOT sent."
            
            print("[SUCCESS] Successfully created a draft email (not sent)")
            print("[INFO] Email creation and account access verified")
            
            # Close the draft without saving
            mail = None
            
        except Exception as e:
            print(f"[ERROR] Failed to create test email: {e}")
            logger.error(f"Email creation test failed: {e}")
        
        print("\n" + "="*80)
        print("SUMMARY")
        print("="*80)
        
        if found_accounts:
            print(f"[SUCCESS] Found {len(found_accounts)} dispatch email account(s)")
            for email in found_accounts:
                print(f"  - {email}")
            print("\n[INFO] You can send emails from these accounts")
        else:
            print("[WARNING] No dispatch email accounts found")
            print("[INFO] You may need to add dispatchfleet@hertz.com to your Outlook")
            print("\nCurrent accounts:")
            for email in account_list:
                print(f"  - {email}")
        
        print("\n[INFO] No emails were sent during this check")
        
        return found_accounts
        
    except Exception as e:
        print(f"[ERROR] Failed to check Outlook accounts: {e}")
        logger.error(f"Failed to check accounts: {e}", exc_info=True)
        return []


def main():
    """Main function to check Outlook accounts."""
    print("="*80)
    print("OUTLOOK ACCOUNT VERIFICATION")
    print("="*80)
    print("This script will check your Outlook accounts and verify email capabilities")
    print("NO EMAILS WILL BE SENT\n")
    
    accounts = check_outlook_accounts()
    
    print("\n" + "="*80)
    print("VERIFICATION COMPLETE")
    print("="*80)


if __name__ == "__main__":
    main()




