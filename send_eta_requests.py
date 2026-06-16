#!/usr/bin/env python3
"""
Send ETA Request Emails to Carriers
"""

from csv_email_automation import CSVEmailAutomation
import os

def main():
    """Send ETA request emails to all carriers."""
    
    # Find CSV file
    csv_files = [f for f in os.listdir('.') if f.endswith('.csv')]
    
    if not csv_files:
        print("❌ No CSV files found in current directory")
        return
    
    csv_file = csv_files[0]
    print(f"📁 Using CSV file: {csv_file}")
    
    try:
        # Create automation instance
        automation = CSVEmailAutomation(csv_file)
        automation.load_csv_data()
        
        print(f"📊 Found {len(automation.orders)} past due orders")
        
        # Show carriers summary
        carriers = automation.get_unique_carriers()
        print(f"🚛 Will email {len(carriers)} unique carriers:")
        
        for carrier in carriers:
            print(f"  • {carrier['name']} ({carrier['order_count']} orders) - {carrier['email']}")
        
        # Show CC information
        cc_email = "joshua.blankenship@hertz.com"
        print(f"\n📧 CC: {cc_email} will be copied on ALL emails")
        
        # Confirm before sending
        print("\n⚠️  READY TO SEND EMAILS")
        print("This will send actual emails to all carriers with Joshua CC'd on every email.")
        
        confirm = input("\nType 'SEND' to confirm and send all emails (or anything else to cancel): ").strip()
        
        if confirm.upper() == 'SEND':
            print("\n📤 Sending emails with CC to Joshua...")
            
            # Send emails with CC (not preview mode)
            results = automation.send_emails_batch(
                preview_mode=False,
                cc_recipients=[cc_email]
            )
            
            # Show results
            print("\n" + "="*50)
            print("📊 EMAIL SENDING RESULTS")
            print("="*50)
            print(f"Total Orders: {results['total_orders']}")
            print(f"✅ Sent Successfully: {results['sent_successfully']}")
            print(f"❌ Failed to Send: {results['failed_to_send']}")
            print(f"⏭️  Skipped: {results['skipped']}")
            
            # Show details
            if results['details']:
                print("\nDETAILS:")
                for detail in results['details']:
                    status_emoji = "✅" if "SUCCESSFULLY" in detail['status'] else "❌" if "FAILED" in detail['status'] or "ERROR" in detail['status'] else "⏭️"
                    print(f"  {status_emoji} {detail['order_id']} - {detail['carrier_name']} ({detail['email']})")
            
            if results['sent_successfully'] > 0:
                print(f"\n🎉 Successfully sent {results['sent_successfully']} ETA request emails!")
                print(f"📧 Joshua ({cc_email}) was CC'd on all emails.")
            
            if results['failed_to_send'] > 0:
                print(f"\n⚠️  {results['failed_to_send']} emails failed to send. Check the details above.")
                
        else:
            print("❌ Email sending cancelled.")
            
    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    main()






