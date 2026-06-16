#!/usr/bin/env python3
"""
Display Draft Emails for All Overdue Orders
Opens each draft in Outlook so you can see and review them
"""

import csv
import win32com.client
from datetime import datetime
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Salesperson email mapping
SALESPERSON_EMAILS = {
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


def parse_date(date_string):
    """Parse date string."""
    if not date_string or date_string.strip() == '':
        return None
    try:
        return datetime.strptime(date_string.strip(), '%m/%d/%Y')
    except ValueError:
        return None


def get_overdue_orders():
    """Get all overdue orders from raw.csv."""
    try:
        with open('raw.csv', 'r', encoding='utf-8') as file:
            csv_reader = csv.DictReader(file)
            all_orders = list(csv_reader)
        
        today = datetime.now()
        
        # Filter for overdue orders
        filtered_orders = []
        for order in all_orders:
            vehicle_status = order.get('Vehicle Status', '').strip()
            carrier_name = order.get('Carrier Name', '').strip()
            delivery_date_str = order.get('Carrier Delivery Scheduled At', '').strip()
            
            if vehicle_status in ['Accepted', 'Picked Up']:
                if carrier_name.lower() != 'oem warranty tows':
                    delivery_date = parse_date(delivery_date_str)
                    if delivery_date and delivery_date.date() < today.date():
                        filtered_orders.append(order)
        
        # Consolidate orders with multiple VINs
        order_groups = {}
        for order in filtered_orders:
            order_id = order.get('Order ID', 'Unknown')
            if order_id not in order_groups:
                order_groups[order_id] = []
            order_groups[order_id].append(order)
        
        overdue_orders = []
        for order_id, order_list in order_groups.items():
            if len(order_list) == 1:
                overdue_orders.append(order_list[0])
            else:
                # Consolidate multiple VINs
                consolidated_order = order_list[0].copy()
                vin_vehicle_pairs = []
                for order in order_list:
                    vin = order.get('VIN #', '').strip()
                    vehicle = order.get('Vehicle Info', '').strip()
                    if vin:
                        vin_vehicle_pairs.append((vin, vehicle))
                
                vins = [pair[0] for pair in vin_vehicle_pairs]
                vehicles = [pair[1] for pair in vin_vehicle_pairs]
                
                consolidated_order['VIN #'] = '; '.join(vins)
                consolidated_order['Vehicle Info'] = '; '.join(vehicles)
                consolidated_order['_multiple_vins'] = True
                consolidated_order['_vin_count'] = len(vins)
                
                overdue_orders.append(consolidated_order)
        
        return overdue_orders
        
    except Exception as e:
        logger.error(f"Failed to load orders: {e}")
        return []


def create_email_body(order_data):
    """Create email body from order data."""
    order_id = order_data.get('Order ID', 'N/A')
    vin = order_data.get('VIN #', 'N/A')
    vehicle_info = order_data.get('Vehicle Info', 'N/A')
    vehicle_status = order_data.get('Vehicle Status', 'N/A')
    carrier_name = order_data.get('Carrier Name', 'N/A')
    
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
        vin_vehicle_lines = []
        for i, (v, veh) in enumerate(zip(vins, vehicles), 1):
            vin_vehicle_lines.append(f"VIN #{i}: {v} - {veh}")
        vin_section = '\n'.join(vin_vehicle_lines)
        vehicle_section = ""
        multiple_note = f" ({len(vins)} vehicles)"
    else:
        vin_section = f"VIN #: {vin}"
        vehicle_section = f"Vehicle: {vehicle_info}"
        multiple_note = ""

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


def display_all_drafts():
    """Display draft emails for all overdue orders."""
    try:
        print("="*80)
        print("DISPLAYING DRAFTS FOR ALL OVERDUE ORDERS")
        print("="*80)
        
        # Get overdue orders
        print("\n[INFO] Loading overdue orders from raw.csv...")
        orders = get_overdue_orders()
        
        if not orders:
            print("[INFO] No overdue orders found.")
            return
        
        print(f"\n[SUCCESS] Found {len(orders)} overdue order(s)")
        
        # Connect to Outlook
        print("\n[INFO] Connecting to Outlook...")
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        # Display drafts for each order
        print(f"\n[INFO] Opening {len(orders)} draft email(s) in Outlook...\n")
        
        for i, order in enumerate(orders, 1):
            order_id = order.get('Order ID', 'N/A')
            carrier_name = order.get('Carrier Name', 'N/A')
            carrier_email = order.get('Carrier Email', 'N/A')
            salesperson = order.get('Salesperson', 'N/A')
            salesperson_email = SALESPERSON_EMAILS.get(salesperson.strip())
            scheduled_delivery = order.get('Carrier Delivery Scheduled At', 'N/A')
            
            print(f"{i}. Opening draft for Order #{order_id}")
            print(f"   Carrier: {carrier_name}")
            print(f"   TO: {carrier_email}")
            if salesperson_email:
                print(f"   CC: {salesperson_email} ({salesperson})")
            print(f"   Delivery Date: {scheduled_delivery}")
            
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
            mail.Subject = f"OVERDUE DELIVERY ALERT - Order #{order_id}"
            mail.Body = create_email_body(order)
            
            # Display the email (opens in Outlook window)
            mail.Display()
            
            print(f"   [SUCCESS] Draft opened in Outlook window #{i}")
            print()
        
        print("="*80)
        print(f"ALL {len(orders)} DRAFT EMAILS OPENED")
        print("="*80)
        
        print(f"\n[SUCCESS] {len(orders)} Outlook windows should now be open")
        print("\n[INFO] You can now:")
        print("  - Review each email")
        print("  - Make any edits")
        print("  - Click 'Send' when ready (or close without sending)")
        print("\n[IMPORTANT] No emails have been sent yet - review and send manually")
        
    except Exception as e:
        print(f"\n[ERROR] Failed to display drafts: {e}")
        logger.error(f"Error: {e}", exc_info=True)


if __name__ == "__main__":
    display_all_drafts()


