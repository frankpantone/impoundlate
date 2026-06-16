#!/usr/bin/env python3
"""
CSV Email Automation for Past Due Orders
Sends ETA request emails to carriers based on CSV data.
"""

import csv
import os
from typing import List, Dict, Optional
from datetime import datetime
from outlook_email_sender import OutlookEmailSender
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class CSVEmailAutomation:
    def __init__(self, csv_file_path: str):
        """
        Initialize CSV email automation.
        
        Args:
            csv_file_path: Path to the CSV file containing order data
        """
        self.csv_file_path = csv_file_path
        self.email_sender = OutlookEmailSender()
        self.orders = []
        
    def load_csv_data(self) -> List[Dict]:
        """
        Load and parse CSV data, grouping by Order ID to handle multiple VINs.
        
        Returns:
            List of dictionaries containing consolidated order data
        """
        try:
            with open(self.csv_file_path, 'r', encoding='utf-8') as file:
                csv_reader = csv.DictReader(file)
                raw_orders = list(csv_reader)
                
            # Group orders by Order ID
            order_groups = {}
            for order in raw_orders:
                order_id = order.get('Order ID', 'Unknown')
                if order_id not in order_groups:
                    order_groups[order_id] = []
                order_groups[order_id].append(order)
            
            # Consolidate orders with multiple VINs
            self.orders = []
            for order_id, order_list in order_groups.items():
                if len(order_list) == 1:
                    # Single VIN - use as is
                    self.orders.append(order_list[0])
                else:
                    # Multiple VINs - consolidate into one order
                    consolidated_order = order_list[0].copy()  # Use first order as base
                    
                    # Collect all VIN-vehicle pairs (preserving order and pairing)
                    vin_vehicle_pairs = []
                    for order in order_list:
                        vin = order.get('VIN #', '').strip()
                        vehicle = order.get('Vehicle Info', '').strip()
                        if vin:  # Only add if VIN exists
                            vin_vehicle_pairs.append((vin, vehicle))
                    
                    # Extract VINs and vehicles maintaining order and pairing
                    vins = [pair[0] for pair in vin_vehicle_pairs]
                    vehicles = [pair[1] for pair in vin_vehicle_pairs]
                    
                    # Update consolidated order with multiple VINs
                    consolidated_order['VIN #'] = '; '.join(vins)
                    consolidated_order['Vehicle Info'] = '; '.join(vehicles)
                    consolidated_order['_multiple_vins'] = True
                    consolidated_order['_vin_count'] = len(vins)
                    
                    self.orders.append(consolidated_order)
            
            logger.info(f"Loaded {len(raw_orders)} raw orders, consolidated to {len(self.orders)} unique orders")
            return self.orders
        except Exception as e:
            logger.error(f"Failed to load CSV data: {e}")
            raise
    
    def create_email_subject(self, order_data: Dict) -> str:
        """
        Create email subject line based on order data.
        
        Args:
            order_data: Dictionary containing order information
            
        Returns:
            Formatted subject line
        """
        order_id = order_data.get('Order ID', 'Unknown')
        return f"ETA Requested - Order ID {order_id}"
    
    def create_email_body(self, order_data: Dict) -> str:
        """
        Create email body based on order data.
        
        Args:
            order_data: Dictionary containing order information
            
        Returns:
            Formatted email body
        """
        # Extract data from CSV columns
        order_id = order_data.get('Order ID', 'N/A')
        vin = order_data.get('VIN #', 'N/A')
        vehicle_info = order_data.get('Vehicle Info', 'N/A')
        
        # Check if this order has multiple VINs
        has_multiple_vins = order_data.get('_multiple_vins', False)
        vin_count = order_data.get('_vin_count', 1)
        
        pickup_business = order_data.get('Pickup Business Name', 'N/A')
        pickup_address = order_data.get('Pickup Address', 'N/A')
        pickup_city = order_data.get('Pickup City', 'N/A')
        pickup_state = order_data.get('Pickup State', 'N/A')
        pickup_zip = order_data.get('Pickup ZIP', 'N/A')
        
        delivery_business = order_data.get('Delivery Business Name', 'N/A')
        delivery_address = order_data.get('Delivery Address', 'N/A')
        delivery_city = order_data.get('Delivery City', 'N/A')
        delivery_state = order_data.get('Delivery State', 'N/A')
        delivery_zip = order_data.get('Delivery ZIP', 'N/A')
        
        carrier_name = order_data.get('Carrier Name', 'N/A')
        scheduled_pickup = order_data.get('Carrier Pickup Scheduled At', 'N/A')
        
        # Format VIN and vehicle information
        if has_multiple_vins:
            vins = vin.split('; ')
            vehicles = vehicle_info.split('; ')
            
            # Consolidate VIN and vehicle on same line
            vin_vehicle_lines = []
            for i, (v, veh) in enumerate(zip(vins, vehicles), 1):
                vin_vehicle_lines.append(f"VIN #{i}: {v} - {veh}")
            
            vin_section = '\n'.join(vin_vehicle_lines)
            vehicle_section = ""  # No separate vehicle section needed
            multiple_note = f" ({len(vins)} vehicles)"
        else:
            vin_section = f"VIN #: {vin}"
            vehicle_section = f"Vehicle: {vehicle_info}"
            multiple_note = ""

        # Create email body
        email_body = f"""Hello {carrier_name},

We hope this email finds you well. We are writing to request an updated ETA for the following past due order{multiple_note}:

Please REPLY ALL to this email at your earliest convenience with your updated timeline or reassignment request.

ORDER DETAILS:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Order ID: {order_id}
Originally Scheduled Pickup: {scheduled_pickup}
{vin_section}
{vehicle_section + chr(10) if vehicle_section else ""}
PICKUP LOCATION:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Business Name: {pickup_business}
Address: {pickup_address}, {pickup_city}, {pickup_state} {pickup_zip}

DELIVERY LOCATION:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Business Name: {delivery_business}
Address: {delivery_address}, {delivery_city}, {delivery_state} {delivery_zip}

ACTION REQUIRED:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
This order is past the scheduled pickup date. Please respond with one of the following:

1. Your updated pickup and delivery ETA
2. If you need this order to be reassigned to another carrier

We understand that delays can occur, and we appreciate your prompt communication to help us keep our customers informed.



Thank you for your continued partnership.

Best regards,
Hertz Transportation Team"""

        return email_body
    
    def create_sample_draft(self, order_index: int = 0) -> Dict:
        """
        Create a sample email draft for user review.
        
        Args:
            order_index: Index of order to use for sample (default: 0)
            
        Returns:
            Dictionary containing sample email data
        """
        if not self.orders:
            self.load_csv_data()
            
        if order_index >= len(self.orders):
            order_index = 0
            
        sample_order = self.orders[order_index]
        
        return {
            'order_data': sample_order,
            'to_email': sample_order.get('Carrier Email', 'N/A'),
            'carrier_name': sample_order.get('Carrier Name', 'N/A'),
            'subject': self.create_email_subject(sample_order),
            'body': self.create_email_body(sample_order)
        }
    
    def send_emails_batch(self, start_index: int = 0, end_index: Optional[int] = None, preview_mode: bool = True, cc_recipients: Optional[List[str]] = None) -> Dict:
        """
        Send emails to carriers in batch.
        
        Args:
            start_index: Starting index for batch sending
            end_index: Ending index for batch sending (None = all remaining)
            preview_mode: If True, only shows what would be sent without actually sending
            cc_recipients: List of email addresses to CC on all emails
            
        Returns:
            Dictionary with sending results
        """
        if not self.orders:
            self.load_csv_data()
            
        if end_index is None:
            end_index = len(self.orders)
            
        results = {
            'total_orders': end_index - start_index,
            'sent_successfully': 0,
            'failed_to_send': 0,
            'skipped': 0,
            'details': []
        }
        
        for i in range(start_index, min(end_index, len(self.orders))):
            order = self.orders[i]
            carrier_email = order.get('Carrier Email', '').strip()
            carrier_name = order.get('Carrier Name', 'Unknown Carrier')
            order_id = order.get('Order ID', 'Unknown')
            
            # Skip if no email address
            if not carrier_email or carrier_email.lower() in ['n/a', 'na', '']:
                results['skipped'] += 1
                results['details'].append({
                    'order_id': order_id,
                    'carrier_name': carrier_name,
                    'status': 'SKIPPED - No email address',
                    'email': carrier_email
                })
                continue
            
            subject = self.create_email_subject(order)
            body = self.create_email_body(order)
            
            if preview_mode:
                results['details'].append({
                    'order_id': order_id,
                    'carrier_name': carrier_name,
                    'status': 'PREVIEW - Would send to',
                    'email': carrier_email,
                    'subject': subject
                })
                continue
            
            # Actual sending
            try:
                success = self.email_sender.send_email(
                    to_recipients=[carrier_email],
                    subject=subject,
                    body=body,
                    cc_recipients=cc_recipients
                )
                
                if success:
                    results['sent_successfully'] += 1
                    status = 'SENT SUCCESSFULLY'
                else:
                    results['failed_to_send'] += 1
                    status = 'FAILED TO SEND'
                    
                results['details'].append({
                    'order_id': order_id,
                    'carrier_name': carrier_name,
                    'status': status,
                    'email': carrier_email
                })
                
            except Exception as e:
                results['failed_to_send'] += 1
                results['details'].append({
                    'order_id': order_id,
                    'carrier_name': carrier_name,
                    'status': f'ERROR: {str(e)}',
                    'email': carrier_email
                })
                logger.error(f"Failed to send email for order {order_id}: {e}")
        
        return results
    
    def get_unique_carriers(self) -> List[Dict]:
        """
        Get list of unique carriers and their email addresses.
        
        Returns:
            List of dictionaries with carrier information
        """
        if not self.orders:
            self.load_csv_data()
            
        carriers = {}
        for order in self.orders:
            carrier_name = order.get('Carrier Name', 'Unknown')
            carrier_email = order.get('Carrier Email', '').strip()
            
            if carrier_name not in carriers:
                carriers[carrier_name] = {
                    'name': carrier_name,
                    'email': carrier_email,
                    'order_count': 0,
                    'orders': []
                }
            
            carriers[carrier_name]['order_count'] += 1
            carriers[carrier_name]['orders'].append(order.get('Order ID', 'Unknown'))
        
        return list(carriers.values())

def main():
    """Main function to demonstrate CSV email automation."""
    # Find CSV file in current directory
    csv_files = [f for f in os.listdir('.') if f.endswith('.csv')]
    
    if not csv_files:
        print("❌ No CSV files found in current directory")
        return
    
    csv_file = csv_files[0]
    print(f"📁 Using CSV file: {csv_file}")
    
    try:
        # Create automation instance
        automation = CSVEmailAutomation(csv_file)
        
        # Load data
        automation.load_csv_data()
        
        print(f"📊 Loaded {len(automation.orders)} orders")
        
        # Show unique carriers
        carriers = automation.get_unique_carriers()
        print(f"🚛 Found {len(carriers)} unique carriers")
        
        # Create sample draft
        sample = automation.create_sample_draft()
        
        print("\n" + "="*60)
        print("📧 SAMPLE EMAIL DRAFT")
        print("="*60)
        print(f"TO: {sample['to_email']}")
        print(f"CARRIER: {sample['carrier_name']}")
        print(f"SUBJECT: {sample['subject']}")
        print("\nBODY:")
        print("-" * 40)
        print(sample['body'])
        print("="*60)
        
        # Preview mode - show what would be sent
        print("\n🔍 PREVIEW MODE - Showing what emails would be sent:")
        results = automation.send_emails_batch(preview_mode=True)
        
        print(f"\nTotal orders: {results['total_orders']}")
        print(f"Would send: {results['total_orders'] - results['skipped']}")
        print(f"Would skip: {results['skipped']} (no email address)")
        
        for detail in results['details'][:5]:  # Show first 5
            print(f"  {detail['status']}: {detail['carrier_name']} ({detail['email']}) - Order {detail['order_id']}")
        
        if len(results['details']) > 5:
            print(f"  ... and {len(results['details']) - 5} more")
            
    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    main()
