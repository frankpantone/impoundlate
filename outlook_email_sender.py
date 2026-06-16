import win32com.client
import sys
from typing import List, Optional
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class OutlookEmailSender:
    def __init__(self):
        """Initialize connection to Outlook application."""
        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            logger.info("Successfully connected to Outlook application")
        except Exception as e:
            logger.error(f"Failed to connect to Outlook: {e}")
            raise
    
    def send_email(self, 
                   to_recipients: List[str],
                   subject: str,
                   body: str,
                   cc_recipients: Optional[List[str]] = None,
                   bcc_recipients: Optional[List[str]] = None,
                   attachments: Optional[List[str]] = None,
                   html_body: bool = False,
                   send_from: Optional[str] = None) -> bool:
        """
        Send an email using Outlook.
        
        Args:
            to_recipients: List of email addresses for TO field
            subject: Email subject line
            body: Email body content
            cc_recipients: List of email addresses for CC field (optional)
            bcc_recipients: List of email addresses for BCC field (optional)
            attachments: List of file paths to attach (optional)
            html_body: Whether the body is HTML formatted (default: False)
            send_from: Email address to send from (for shared mailboxes) (optional)
            
        Returns:
            bool: True if email was sent successfully, False otherwise
        """
        try:
            # Create a new mail item
            mail = self.outlook.CreateItem(0)  # 0 = olMailItem
            
            # Set the sending account if specified (for shared mailboxes)
            if send_from:
                try:
                    namespace = self.outlook.GetNamespace("MAPI")
                    accounts = namespace.Accounts
                    
                    # First try to find exact account match
                    account_found = False
                    for i in range(1, accounts.Count + 1):
                        account = accounts.Item(i)
                        if account.SmtpAddress.lower() == send_from.lower():
                            mail._oleobj_.Invoke(*(64209, 0, 8, 0, account))
                            logger.info(f"Set sending account to: {send_from}")
                            account_found = True
                            break
                    
                    # If not found in accounts, try setting SentOnBehalfOfName for shared mailbox
                    if not account_found:
                        # For shared mailboxes, we need to get the folder
                        folders = namespace.Folders
                        for i in range(1, folders.Count + 1):
                            folder = folders.Item(i)
                            folder_name = folder.Name.lower()
                            # Check if folder name matches or contains the email
                            if send_from.lower() in folder_name or 'dispatch' in folder_name:
                                try:
                                    # Try to get the email address from store
                                    store = folder.Store
                                    # Set the SendUsingAccount property
                                    mail.SendUsingAccount = account  # Use primary account
                                    # But send from the shared mailbox
                                    mail.SentOnBehalfOfName = send_from
                                    logger.info(f"Set shared mailbox sender to: {send_from}")
                                    account_found = True
                                    break
                                except:
                                    pass
                        
                        if not account_found:
                            # Last resort: just set SentOnBehalfOfName
                            mail.SentOnBehalfOfName = send_from
                            logger.info(f"Set 'Send on behalf of' to: {send_from}")
                
                except Exception as e:
                    logger.warning(f"Could not set sending account to {send_from}: {e}")
                    logger.info("Email will be sent from default account")
            
            # Set recipients
            mail.To = "; ".join(to_recipients)
            
            if cc_recipients:
                mail.CC = "; ".join(cc_recipients)
                
            if bcc_recipients:
                mail.BCC = "; ".join(bcc_recipients)
            
            # Set subject and body
            mail.Subject = subject
            
            if html_body:
                mail.HTMLBody = body
            else:
                mail.Body = body
            
            # Add attachments if provided
            if attachments:
                for attachment_path in attachments:
                    try:
                        mail.Attachments.Add(attachment_path)
                        logger.info(f"Added attachment: {attachment_path}")
                    except Exception as e:
                        logger.warning(f"Failed to add attachment {attachment_path}: {e}")
            
            # Send the email
            mail.Send()
            logger.info(f"Email sent successfully to: {', '.join(to_recipients)}")
            return True
            
        except Exception as e:
            logger.error(f"Failed to send email: {e}")
            return False
    
    def send_test_email(self, test_recipient: str) -> bool:
        """
        Send a test email to verify the setup is working.
        
        Args:
            test_recipient: Email address to send test email to
            
        Returns:
            bool: True if test email was sent successfully
        """
        subject = "Test Email from Python Outlook Automation"
        body = """This is a test email sent from Python using Outlook COM interface.

If you received this email, the automation setup is working correctly!

Best regards,
Your Python Email Automation Script"""
        
        return self.send_email(
            to_recipients=[test_recipient],
            subject=subject,
            body=body
        )
    
    def get_current_user_email(self) -> str:
        """
        Get the email address of the current Outlook user.
        
        Returns:
            str: Current user's email address
        """
        try:
            namespace = self.outlook.GetNamespace("MAPI")
            accounts = namespace.Accounts
            if accounts.Count > 0:
                # Get the first account's email address
                return accounts.Item(1).SmtpAddress
            else:
                logger.warning("No accounts found in Outlook")
                return ""
        except Exception as e:
            logger.error(f"Failed to get current user email: {e}")
            return ""

def main():
    """Example usage of the OutlookEmailSender class."""
    try:
        # Create email sender instance
        email_sender = OutlookEmailSender()
        
        # Get current user's email for reference
        current_email = email_sender.get_current_user_email()
        if current_email:
            print(f"Connected to Outlook account: {current_email}")
        
        # Example: Send a test email
        # Replace with your actual email address for testing
        test_email = input("Enter your email address to send a test email: ").strip()
        
        if test_email:
            print("Sending test email...")
            if email_sender.send_test_email(test_email):
                print("✅ Test email sent successfully!")
            else:
                print("❌ Failed to send test email")
        
        # Example: Send a custom email
        print("\n" + "="*50)
        print("Example: Sending a custom email")
        
        recipients = ["example@domain.com"]  # Replace with actual recipients
        subject = "Custom Email from Python"
        body = "This is a custom email sent from Python!"
        
        # Uncomment the lines below to send a custom email
        # if email_sender.send_email(recipients, subject, body):
        #     print("✅ Custom email sent successfully!")
        # else:
        #     print("❌ Failed to send custom email")
        
    except Exception as e:
        logger.error(f"Application error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()







