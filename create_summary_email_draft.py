#!/usr/bin/env python3
"""
Create Summary Email Draft in Outlook
Opens the summary email as a draft in Outlook for review
"""

from send_overdue_delivery_emails import OverdueDeliveryNotifier
import os
import win32com.client
from datetime import datetime
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


def main():
    """Create summary email draft in Outlook."""
    
    print("="*80)
    print("CREATING SUMMARY EMAIL DRAFT IN OUTLOOK")
    print("="*80)
    
    csv_file = 'raw.csv'
    
    if not os.path.exists(csv_file):
        print(f"[ERROR] CSV file '{csv_file}' not found")
        return
    
    try:
        # Create notifier instance
        print(f"\n[INFO] Loading data from {csv_file}...")
        notifier = OverdueDeliveryNotifier(csv_file)
        notifier.load_and_filter_orders()
        
        if not notifier.overdue_orders:
            print("[INFO] No overdue orders found")
            return
        
        print(f"[INFO] Found {len(notifier.overdue_orders)} overdue order(s)")
        
        # Create mock results
        results = {
            'total_orders': len(notifier.overdue_orders),
            'sent_successfully': 0,
            'failed_to_send': 0,
            'skipped': 0,
            'details': []
        }
        
        # Add details for each order
        for order in notifier.overdue_orders:
            order_id = order.get('Order ID', 'Unknown')
            carrier_email = order.get('Carrier Email', '').strip()
            carrier_name = order.get('Carrier Name', 'Unknown Carrier')
            salesperson = order.get('Salesperson', '').strip()
            salesperson_email = notifier.get_salesperson_email(salesperson)
            
            if not carrier_email:
                status = 'SKIPPED - No email'
                results['skipped'] += 1
            else:
                status = 'SENT SUCCESSFULLY'
                results['sent_successfully'] += 1
            
            results['details'].append({
                'order_id': order_id,
                'carrier_name': carrier_name,
                'status': status,
                'email': carrier_email,
                'salesperson': salesperson,
                'cc': salesperson_email
            })
        
        # Create Excel file
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        excel_file = f'DRAFT_summary_report_{timestamp}.xlsx'
        
        print(f"[INFO] Generating Excel report...")
        report_file = notifier.export_to_excel(results, filename=excel_file)
        
        if not report_file:
            print("[ERROR] Failed to create Excel report")
            return
        
        print(f"[SUCCESS] Excel report created: {excel_file}")
        
        # Connect to Outlook
        print(f"\n[INFO] Connecting to Outlook...")
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        # Create the summary email
        print(f"[INFO] Creating summary email draft...")
        mail = outlook.CreateItem(0)
        
        # Set sender to dispatch@hertz.com
        folders = namespace.Folders
        for i in range(1, folders.Count + 1):
            folder = folders.Item(i)
            if 'dispatch' in folder.Name.lower():
                mail.SentOnBehalfOfName = 'dispatch@hertz.com'
                print(f"[SUCCESS] Set sender to: dispatch@hertz.com")
                break
        
        # Set recipients (all salespeople)
        team_emails = notifier.TEAM_EMAILS
        mail.To = "; ".join(team_emails)
        
        # Set subject
        mail.Subject = f"Overdue Delivery Notifications Sent - {datetime.now().strftime('%m/%d/%Y')}"
        
        # Count by salesperson
        salesperson_counts = {}
        for detail in results['details']:
            sp = detail.get('salesperson', 'Unknown')
            if sp not in salesperson_counts:
                salesperson_counts[sp] = {'sent': 0, 'failed': 0, 'skipped': 0}
            
            if 'SUCCESSFULLY' in detail['status']:
                salesperson_counts[sp]['sent'] += 1
            elif 'FAILED' in detail['status'] or 'ERROR' in detail['status']:
                salesperson_counts[sp]['failed'] += 1
            else:
                salesperson_counts[sp]['skipped'] += 1
        
        # Create email body
        timestamp_text = datetime.now().strftime('%B %d, %Y at %I:%M %p')
        
        body = f"""Team,

Overdue delivery notification emails have been processed on {timestamp_text}.

SUMMARY:
Total Orders Processed: {results['total_orders']}
Emails Sent Successfully: {results['sent_successfully']}
Failed to Send: {results['failed_to_send']}
Skipped: {results['skipped']}

BREAKDOWN BY SALESPERSON:

"""
        
        for sp, counts in sorted(salesperson_counts.items()):
            total = counts['sent'] + counts['failed'] + counts['skipped']
            if total > 0:
                body += f"{sp}: {total}\n"
        
        body += f"""
ATTACHED:
The attached Excel file contains complete details of all overdue orders processed.

Please review the attached report for full details on your assigned orders.

Best regards,
Hertz Dispatch Fleet Team
dispatch@hertz.com"""
        
        mail.Body = body
        
        # Attach Excel file
        full_path = os.path.abspath(excel_file)
        mail.Attachments.Add(full_path)
        print(f"[INFO] Attached: {excel_file}")
        
        # Display the email
        print(f"\n[INFO] Opening summary email draft in Outlook...")
        mail.Display()
        
        print("\n" + "="*80)
        print("SUMMARY EMAIL DRAFT OPENED IN OUTLOOK")
        print("="*80)
        
        print(f"\nFROM: dispatch@hertz.com")
        print(f"TO: {len(team_emails)} salespeople")
        for email in team_emails:
            print(f"  - {email}")
        
        print(f"\nSUBJECT: {mail.Subject}")
        print(f"ATTACHMENT: {excel_file}")
        
        print(f"\n[SUCCESS] Summary email draft is now open in Outlook!")
        print(f"[INFO] You can:")
        print(f"  - Review the email content")
        print(f"  - Review the attached Excel file")
        print(f"  - Close without sending")
        print(f"  - Make any edits")
        
        print(f"\n[IMPORTANT] This is a draft for review - NOT sent")
        
    except Exception as e:
        print(f"[ERROR] Error: {e}")
        logger.error(f"Application error: {e}", exc_info=True)


if __name__ == "__main__":
    main()

