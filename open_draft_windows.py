#!/usr/bin/env python3
"""
Open Draft Email Windows for Review
Opens draft emails in Outlook windows instead of just saving to Drafts
"""

from send_overdue_delivery_emails import OverdueDeliveryNotifier
import os
import win32com.client
import logging

# Note: Filtering is handled by OverdueDeliveryNotifier.load_and_filter_orders()
# which now excludes 'OEM Warranty tows' (case-insensitive)

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


def main():
    """Create and display draft emails."""
    
    print("="*80)
    print("OPENING DRAFT EMAIL WINDOWS FOR REVIEW")
    print("="*80)
    
    csv_file = 'raw.csv'
    
    if not os.path.exists(csv_file):
        print(f"[ERROR] CSV file '{csv_file}' not found")
        return
    
    print(f"\n[INFO] Using CSV file: {csv_file}")
    
    try:
        # Create notifier instance
        notifier = OverdueDeliveryNotifier(csv_file)
        notifier.load_and_filter_orders()
        
        if not notifier.overdue_orders:
            print("[SUCCESS] No overdue deliveries found!")
            return
        
        print(f"[INFO] Found {len(notifier.overdue_orders)} overdue order(s)")
        
        # Ask how many to open
        print(f"\n[INFO] Opening all {len(notifier.overdue_orders)} emails in separate windows...")
        print("[INFO] Each email will open in its own Outlook window for review")
        
        confirm = input(f"\nOpen {len(notifier.overdue_orders)} draft windows? (y/n): ").strip().lower()
        
        if confirm != 'y':
            print("[CANCELLED] Draft creation cancelled")
            return
        
        # Connect to Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        print(f"\n[INFO] Opening {len(notifier.overdue_orders)} draft email windows...")
        print()
        
        for i, order in enumerate(notifier.overdue_orders, 1):
            order_id = order.get('Order ID', 'N/A')
            carrier_name = order.get('Carrier Name', 'N/A')
            carrier_email = order.get('Carrier Email', 'N/A')
            salesperson = order.get('Salesperson', 'N/A')
            salesperson_email = notifier.get_salesperson_email(salesperson)
            vin_count = order.get('_vin_count', 1)
            
            if not carrier_email or carrier_email.strip() == '':
                print(f"{i}. [SKIPPED] Order #{order_id} - No carrier email")
                continue
            
            print(f"{i}. Opening: Order #{order_id} - {carrier_name} ({vin_count} vehicle{'s' if vin_count > 1 else ''})")
            
            # Create mail item
            mail = outlook.CreateItem(0)
            
            # Set sender to dispatch@hertz.com
            folders = namespace.Folders
            for j in range(1, folders.Count + 1):
                folder = folders.Item(j)
                if 'dispatch' in folder.Name.lower():
                    mail.SentOnBehalfOfName = 'dispatch@hertz.com'
                    break
            
            # Set recipients
            mail.To = carrier_email
            if salesperson_email:
                mail.CC = salesperson_email
            
            # Set subject and body
            mail.Subject = notifier.create_email_subject(order)
            mail.Body = notifier.create_email_body(order)
            
            # Display the email (opens window)
            mail.Display()
        
        print(f"\n[SUCCESS] Opened {len(notifier.overdue_orders)} draft email windows")
        print(f"\n[INFO] You should now see {len(notifier.overdue_orders)} Outlook windows")
        print("[INFO] Review each email and:")
        print("  - Click 'Send' to send the email")
        print("  - Close without saving to discard")
        print("  - Make edits and then send")
        
    except Exception as e:
        print(f"[ERROR] Error: {e}")
        logger.error(f"Application error: {e}", exc_info=True)


if __name__ == "__main__":
    main()

