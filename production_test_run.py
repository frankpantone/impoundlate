#!/usr/bin/env python3
"""
Production Test Run - Create Drafts and Excel Report
Processes raw.csv and creates draft emails + Excel output WITHOUT sending
"""

from send_overdue_delivery_emails import OverdueDeliveryNotifier
import os
import win32com.client
from datetime import datetime
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


def create_draft_emails(notifier):
    """Create draft emails for all overdue orders."""
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        print("\n[INFO] Creating draft emails in Outlook...")
        print("[INFO] These will be saved to your Drafts folder for review\n")
        
        created_drafts = []
        
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
            
            print(f"{i}. Creating draft for Order #{order_id}")
            print(f"   Carrier: {carrier_name}")
            print(f"   Vehicles: {vin_count}")
            
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
            
            # Save as draft
            mail.Save()
            
            created_drafts.append({
                'order_id': order_id,
                'carrier': carrier_name,
                'salesperson': salesperson,
                'vin_count': vin_count
            })
            
            print(f"   [SUCCESS] Draft saved to Drafts folder")
        
        return created_drafts
        
    except Exception as e:
        logger.error(f"Failed to create drafts: {e}")
        return []


def main():
    """Run production test without sending emails."""
    
    print("="*80)
    print("PRODUCTION TEST RUN - DRAFTS & EXCEL REPORT")
    print("="*80)
    print("\n[INFO] This test will:")
    print("  1. Process raw.csv file")
    print("  2. Filter for overdue orders")
    print("  3. Create draft emails (saved to Outlook Drafts)")
    print("  4. Generate Excel output report")
    print("  5. NOT send any emails")
    print("\n" + "="*80)
    
    csv_file = 'raw.csv'
    
    if not os.path.exists(csv_file):
        print(f"[ERROR] CSV file '{csv_file}' not found")
        return
    
    print(f"\n[INFO] Using CSV file: {csv_file}")
    
    try:
        # Create notifier instance
        notifier = OverdueDeliveryNotifier(csv_file)
        
        # Load and filter orders
        print("[INFO] Scanning for overdue deliveries...")
        notifier.load_and_filter_orders()
        
        if not notifier.overdue_orders:
            print("[SUCCESS] No overdue deliveries found! All orders are on schedule.")
            return
        
        print(f"\n[INFO] Found {len(notifier.overdue_orders)} overdue order(s)")
        
        # Show preview
        print("\n" + "="*80)
        print("OVERDUE ORDERS SUMMARY")
        print("="*80)
        
        for i, order in enumerate(notifier.overdue_orders, 1):
            order_id = order.get('Order ID', 'N/A')
            carrier_name = order.get('Carrier Name', 'N/A')
            salesperson = order.get('Salesperson', 'N/A')
            delivery_date = order.get('Carrier Delivery Scheduled At', 'N/A')
            vin_count = order.get('_vin_count', 1)
            
            print(f"\n{i}. Order #{order_id}")
            print(f"   Salesperson: {salesperson}")
            print(f"   Carrier: {carrier_name}")
            print(f"   Scheduled Delivery: {delivery_date}")
            print(f"   Vehicles: {vin_count}")
        
        # Create drafts
        print("\n" + "="*80)
        print("CREATING DRAFT EMAILS")
        print("="*80)
        
        drafts = create_draft_emails(notifier)
        
        print(f"\n[SUCCESS] Created {len(drafts)} draft email(s)")
        print("[INFO] Drafts saved to: Outlook > Drafts folder")
        
        # Simulate email results for Excel export
        print("\n" + "="*80)
        print("GENERATING EXCEL REPORT")
        print("="*80)
        
        # Create results dictionary for Excel export
        results = {
            'total_orders': len(notifier.overdue_orders),
            'sent_successfully': 0,  # None sent in test mode
            'failed_to_send': 0,
            'skipped': 0,
            'details': []
        }
        
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
                status = 'DRAFT CREATED (Preview Mode)'
            
            results['details'].append({
                'order_id': order_id,
                'carrier_name': carrier_name,
                'status': status,
                'email': carrier_email,
                'salesperson': salesperson,
                'cc': salesperson_email
            })
        
        # Generate Excel report
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        excel_file = f'TEST_RUN_overdue_delivery_emails_{timestamp}.xlsx'
        
        print(f"[INFO] Creating Excel report: {excel_file}")
        
        report_file = notifier.export_to_excel(results, filename=excel_file)
        
        if report_file:
            full_path = os.path.abspath(report_file)
            
            # Preview summary email (don't actually send during test)
            print("\n" + "="*80)
            print("PREVIEWING SUMMARY EMAIL TO SALESPEOPLE")
            print("="*80)
            notifier.send_summary_email(report_file, results, send_email=False)
            
            print("\n" + "="*80)
            print("PRODUCTION TEST COMPLETE")
            print("="*80)
            
            print(f"\n[SUCCESS] {len(drafts)} draft email(s) created")
            print(f"[SUCCESS] Excel report generated")
            
            print(f"\n[DRAFTS] LOCATION:")
            print(f"   Open Outlook > Drafts folder")
            print(f"   You will find {len(drafts)} draft emails to review")
            
            print(f"\n[EXCEL] REPORT:")
            print(f"   File: {excel_file}")
            print(f"   Path: {full_path}")
            
            print(f"\n[REPORT] CONTAINS:")
            print(f"   - {len(notifier.overdue_orders)} overdue order(s)")
            print(f"   - Salesperson assignments")
            print(f"   - Pickup & Delivery details")
            print(f"   - VIN counts")
            print(f"   - Email status (Draft Created)")
            
            print(f"\n[SUMMARY] EMAIL (PREVIEW ABOVE):")
            print(f"   - Will be sent to all 9 salespeople")
            print(f"   - Includes Excel report as attachment")
            print(f"   - Sent FROM: dispatch@hertz.com")
            print(f"   - Contains breakdown by salesperson")
            
            print(f"\n[NEXT] STEPS:")
            print(f"   1. Open the Excel file to review all orders")
            print(f"   2. Open Outlook Drafts to review email content")
            print(f"   3. When ready to send, run: python send_overdue_delivery_emails.py")
            
            print(f"\n[IMPORTANT]:")
            print(f"   - NO emails have been sent (test mode)")
            print(f"   - All drafts are in your Outlook Drafts folder")
            print(f"   - Summary email was NOT sent (preview only)")
            print(f"   - You can review, edit, or delete drafts")
            
        else:
            print("[ERROR] Failed to create Excel report")
        
    except Exception as e:
        print(f"[ERROR] Error: {e}")
        logger.error(f"Application error: {e}", exc_info=True)


if __name__ == "__main__":
    main()



