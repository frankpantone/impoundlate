#!/usr/bin/env python3
"""
Preview Summary Email to Sales Team
Shows what the summary email will look like WITHOUT sending
"""

from send_overdue_delivery_emails import OverdueDeliveryNotifier
import os
from datetime import datetime
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


def main():
    """Preview the summary email."""
    
    print("="*80)
    print("PREVIEW: SUMMARY EMAIL TO SALES TEAM")
    print("="*80)
    
    csv_file = 'raw.csv'
    
    if not os.path.exists(csv_file):
        print(f"[ERROR] CSV file '{csv_file}' not found")
        return
    
    try:
        # Create notifier instance
        notifier = OverdueDeliveryNotifier(csv_file)
        notifier.load_and_filter_orders()
        
        if not notifier.overdue_orders:
            print("[INFO] No overdue orders found")
            return
        
        print(f"\n[INFO] Found {len(notifier.overdue_orders)} overdue order(s)")
        
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
        
        # Create Excel file for preview
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        excel_file = f'PREVIEW_summary_report_{timestamp}.xlsx'
        
        print(f"[INFO] Generating sample Excel report...")
        report_file = notifier.export_to_excel(results, filename=excel_file)
        
        if report_file:
            print(f"[SUCCESS] Sample Excel created: {excel_file}")
            
            # Show the summary email preview
            print("\n" + "="*80)
            print("SUMMARY EMAIL PREVIEW")
            print("="*80)
            
            # Get recipient list
            team_emails = notifier.TEAM_EMAILS
            
            print(f"\nFROM: dispatch@hertz.com")
            print(f"\nTO ({len(team_emails)} salespeople):")
            for email in team_emails:
                print(f"  - {email}")
            
            print(f"\nSUBJECT: Overdue Delivery Notifications Sent - {datetime.now().strftime('%m/%d/%Y')}")
            
            print(f"\nATTACHMENT: {excel_file}")
            
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
            
            timestamp_text = datetime.now().strftime('%B %d, %Y at %I:%M %p')
            
            # Show email body
            print("\nBODY:")
            print("-" * 80)
            print(f"""Team,

Overdue delivery notification emails have been processed on {timestamp_text}.

SUMMARY:
Total Orders Processed: {results['total_orders']}
Emails Sent Successfully: {results['sent_successfully']}
Failed to Send: {results['failed_to_send']}
Skipped: {results['skipped']}

BREAKDOWN BY SALESPERSON:
""")
            
            for sp, counts in sorted(salesperson_counts.items()):
                total = counts['sent'] + counts['failed'] + counts['skipped']
                if total > 0:
                    print(f"{sp}: {total}")
            
            print(f"""
ATTACHED:
The attached Excel file contains complete details of all overdue orders processed.

Please review the attached report for full details on your assigned orders.

Best regards,
Hertz Dispatch Fleet Team
dispatch@hertz.com""")
            
            print("-" * 80)
            
            print("\n" + "="*80)
            print("SUMMARY EMAIL DETAILS")
            print("="*80)
            print(f"\n[INFO] This email will be sent to:")
            print(f"  Total Recipients: {len(team_emails)}")
            print(f"  Attachment Size: {os.path.getsize(report_file) / 1024:.1f} KB")
            print(f"  Attachment File: {excel_file}")
            
            print(f"\n[IMPORTANT] During production run:")
            print(f"  - This email WILL be automatically sent to all salespeople")
            print(f"  - They will receive the Excel report as an attachment")
            print(f"  - They can filter the report to see their assigned orders")
            
            print(f"\n[TEST MODE] Right now:")
            print(f"  - NO emails have been sent")
            print(f"  - This is a preview only")
            
        else:
            print("[ERROR] Failed to create sample Excel report")
        
    except Exception as e:
        print(f"[ERROR] Error: {e}")
        logger.error(f"Application error: {e}", exc_info=True)


if __name__ == "__main__":
    main()

