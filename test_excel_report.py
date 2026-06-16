#!/usr/bin/env python3
"""
Test Excel Report Generation
Creates a sample Excel report to show what the output will look like
"""

from send_overdue_delivery_emails import OverdueDeliveryNotifier
import os

def main():
    """Generate a test Excel report with preview data."""
    
    print("="*80)
    print("TESTING EXCEL REPORT GENERATION")
    print("="*80)
    
    # Create notifier instance
    csv_file = 'raw.csv'
    
    if not os.path.exists(csv_file):
        print(f"[ERROR] CSV file '{csv_file}' not found")
        return
    
    print(f"\n[INFO] Using CSV file: {csv_file}")
    
    try:
        # Create notifier and load data
        notifier = OverdueDeliveryNotifier(csv_file)
        notifier.load_and_filter_orders()
        
        print(f"[INFO] Found {len(notifier.overdue_orders)} overdue orders")
        
        # Create preview results (simulating what would happen after sending)
        results = notifier.send_emails(preview_mode=True)
        
        print(f"\n[INFO] Generating Excel report with {results['total_orders']} orders...")
        
        # Generate Excel file
        excel_file = notifier.export_to_excel(results, filename='TEST_overdue_delivery_report.xlsx')
        
        if excel_file:
            full_path = os.path.abspath(excel_file)
            print("\n" + "="*80)
            print("[SUCCESS] TEST EXCEL REPORT CREATED")
            print("="*80)
            print(f"\nFile: {excel_file}")
            print(f"Full Path: {full_path}")
            print("\n[INFO] The Excel file contains:")
            print("  - Order ID")
            print("  - Status (Accepted/Picked Up)")
            print("  - Carrier Name & Email")
            print("  - Salesperson & CC Email")
            print("  - Scheduled Delivery Date")
            print("  - Days Overdue")
            print("  - VIN Count (for multi-vehicle orders)")
            print("  - Email Status (color-coded)")
            print("  - Summary statistics at bottom")
            
            print("\n[INFO] Color Coding:")
            print("  - Green = Successfully Sent")
            print("  - Red = Failed")
            print("  - Yellow = Skipped")
            
            print(f"\n[INFO] Open this file in Excel to review the format")
            print(f"[INFO] When you run the main script and send emails, a similar")
            print(f"       report will be automatically generated with actual results")
            
        else:
            print("[ERROR] Failed to create Excel report")
            
    except Exception as e:
        print(f"[ERROR] Error: {e}")


if __name__ == "__main__":
    main()

