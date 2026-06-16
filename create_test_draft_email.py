#!/usr/bin/env python3
"""
Create Test Draft Email from Real Raw.csv Data
Creates a draft email using actual overdue order data - DOES NOT SEND
"""

import csv
import win32com.client
from datetime import datetime
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


def parse_date(date_string):
    """Parse date string."""
    if not date_string or date_string.strip() == '':
        return None
    try:
        return datetime.strptime(date_string.strip(), '%m/%d/%Y')
    except ValueError:
        return None


def get_overdue_order():
    """Get one overdue order from raw.csv."""
    try:
        with open('raw.csv', 'r', encoding='utf-8') as file:
            csv_reader = csv.DictReader(file)
            all_orders = list(csv_reader)
        
        today = datetime.now()
        
        # Find first overdue order
        for order in all_orders:
            vehicle_status = order.get('Vehicle Status', '').strip()
            carrier_name = order.get('Carrier Name', '').strip()
            delivery_date_str = order.get('Carrier Delivery Scheduled At', '').strip()
            
            if vehicle_status in ['Accepted', 'Picked Up']:
                if carrier_name != 'OEM Warranty tows':
                    delivery_date = parse_date(delivery_date_str)
                    if delivery_date and delivery_date.date() < today.date():
                        return order
        
        return None
        
    except Exception as e:
        logger.error(f"Failed to load order: {e}")
        return None


def create_email_body(order_data):
    """Create email body from order data."""
    order_id = order_data.get('Order ID', 'N/A')
    vin = order_data.get('VIN #', 'N/A')
    vehicle_info = order_data.get('Vehicle Info', 'N/A')
    vehicle_status = order_data.get('Vehicle Status', 'N/A')
    carrier_name = order_data.get('Carrier Name', 'N/A')
    
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
    
    scheduled_delivery = order_data.get('Carrier Delivery Scheduled At', 'N/A')
    scheduled_pickup = order_data.get('Carrier Pickup Scheduled At', 'N/A')
    
    # Calculate days overdue
    delivery_date = parse_date(scheduled_delivery)
    if delivery_date:
        today = datetime.now()
        days_overdue = (today.date() - delivery_date.date()).days
        # Only show overdue text if actually overdue (positive days)
        if days_overdue > 0:
            overdue_text = f" ({days_overdue} day{'s' if days_overdue != 1 else ''} overdue)"
        else:
            overdue_text = ""
    else:
        overdue_text = ""
    
    # Format VIN and vehicle information
    if has_multiple_vins:
        vins = vin.split('; ')
        vehicles = vehicle_info.split('; ')
        
        # Consolidate VIN and vehicle on same line (matching send_eta_requests format)
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

We hope this email finds you well. We are writing to notify you that the following order{multiple_note} is past the scheduled delivery date.

Please REPLY ALL to this email at your earliest convenience with an updated delivery ETA.

ORDER DETAILS:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Order ID: {order_id}
Current Status: {vehicle_status}
Originally Scheduled Delivery: {scheduled_delivery}{overdue_text}
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
This order is past the scheduled delivery date. Please respond with one of the following:

1. Your updated delivery ETA
2. Current location and status of the vehicle{"s" if has_multiple_vins else ""}
3. If you need this order to be reassigned to another carrier

We understand that delays can occur, and we appreciate your prompt communication to help us keep our customers informed.



Thank you for your continued partnership.

Best regards,
Hertz Dispatch Fleet Team
dispatch@hertz.com"""

    return email_body


def get_salesperson_email(salesperson_name):
    """Get salesperson email."""
    salesperson_emails = {
        'Geo Marrero': 'geo.marrero@hertz.com',
        'Alexei Piljan': 'alexei.piljan@hertz.com',
        'Patrick Gal': 'patrick.gal@hertz.com',
        'Molly Sutphen': 'Molly.Sutphen@hertz.com',
        'Enrique Martinez': 'E.Martinez@hertz.com',
        'Trung Thai': 'trung.thai@hertz.com',
        'Ryan Cappelen': 'Ryan.Cappelen@hertz.com',
        'Matt Biocchi': 'matthew.biocchi@hertz.com',
        'Josh Blankenship': 'joshua.blankenship@hertz.com'
    }
    return salesperson_emails.get(salesperson_name.strip())


def create_test_draft():
    """Create a test draft email from real data."""
    try:
        print("="*80)
        print("CREATING TEST DRAFT EMAIL FROM REAL DATA")
        print("="*80)
        
        # Get overdue order
        print("\n[INFO] Loading overdue order from raw.csv...")
        order = get_overdue_order()
        
        if not order:
            print("[ERROR] No overdue orders found in raw.csv")
            return
        
        order_id = order.get('Order ID', 'N/A')
        carrier_name = order.get('Carrier Name', 'N/A')
        carrier_email = order.get('Carrier Email', 'N/A')
        salesperson = order.get('Salesperson', 'N/A')
        salesperson_email = get_salesperson_email(salesperson)
        
        print(f"\n[INFO] Using Order: {order_id}")
        print(f"  Carrier: {carrier_name}")
        print(f"  Carrier Email: {carrier_email}")
        print(f"  Salesperson: {salesperson}")
        if salesperson_email:
            print(f"  Salesperson Email: {salesperson_email}")
        
        # Connect to Outlook
        print("\n[INFO] Connecting to Outlook...")
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        # Create mail item
        print("[INFO] Creating draft email...")
        mail = outlook.CreateItem(0)
        
        # Set sender to dispatch@hertz.com
        folders = namespace.Folders
        for i in range(1, folders.Count + 1):
            folder = folders.Item(i)
            if 'dispatch' in folder.Name.lower():
                mail.SentOnBehalfOfName = 'dispatch@hertz.com'
                print(f"[SUCCESS] Set sender to: dispatch@hertz.com")
                break
        
        # Set recipients
        mail.To = carrier_email
        if salesperson_email:
            mail.CC = salesperson_email
            print(f"[INFO] CC: {salesperson_email}")
        
        # Set subject
        mail.Subject = f"OVERDUE DELIVERY ALERT - Order #{order_id}"
        
        # Set body
        mail.Body = create_email_body(order)
        
        # Display the email (opens it in a window)
        print("\n[INFO] Opening draft email for your review...")
        mail.Display()  # This opens the email in Outlook so you can see it
        
        print("\n" + "="*80)
        print("DRAFT EMAIL OPENED IN OUTLOOK")
        print("="*80)
        print(f"\nFROM: dispatch@hertz.com")
        print(f"TO: {carrier_email}")
        if salesperson_email:
            print(f"CC: {salesperson_email}")
        print(f"SUBJECT: OVERDUE DELIVERY ALERT - Order #{order_id}")
        print("\n[SUCCESS] Draft email is now open in Outlook for your review")
        print("[INFO] You can:")
        print("  - Review the email content")
        print("  - Close it without saving")
        print("  - Save it to drafts")
        print("  - Make any edits you want")
        print("\n[IMPORTANT] This email is NOT sent - it's only a draft for review")
        
    except Exception as e:
        print(f"\n[ERROR] Failed to create draft: {e}")
        logger.error(f"Error: {e}", exc_info=True)


if __name__ == "__main__":
    create_test_draft()

