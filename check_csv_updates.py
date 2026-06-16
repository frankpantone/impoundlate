import csv
from datetime import datetime

with open('raw.csv', 'r', encoding='utf-8') as f:
    rows = list(csv.DictReader(f))

print(f"Total rows in CSV: {len(rows)}")
print("\n" + "="*80)
print("LAST 10 ROWS IN CSV:")
print("="*80)

for i, row in enumerate(rows[-10:], len(rows)-9):
    order_id = row.get('Order ID', 'N/A')
    status = row.get('Vehicle Status', 'N/A')
    delivery = row.get('Carrier Delivery Scheduled At', 'N/A')
    carrier = row.get('Carrier Name', 'N/A')
    vin = row.get('VIN #', 'N/A')
    
    print(f"\n{i}. Order: {order_id}")
    print(f"   Status: {status}")
    print(f"   Delivery Date: {delivery}")
    print(f"   Carrier: {carrier}")
    print(f"   VIN: {vin[:20]}...")

# Check for orders with same Order ID (multiple VINs)
print("\n" + "="*80)
print("ORDERS WITH MULTIPLE VINS:")
print("="*80)

order_groups = {}
for row in rows:
    order_id = row.get('Order ID', '')
    if order_id not in order_groups:
        order_groups[order_id] = []
    order_groups[order_id].append(row)

multi_vin_orders = {oid: orders for oid, orders in order_groups.items() if len(orders) > 1}

print(f"Found {len(multi_vin_orders)} orders with multiple VINs:")
for oid, orders in sorted(multi_vin_orders.items()):
    print(f"\n  {oid}: {len(orders)} VINs")
    print(f"    Status: {orders[0].get('Vehicle Status', 'N/A')}")
    print(f"    Delivery: {orders[0].get('Carrier Delivery Scheduled At', 'N/A')}")
    print(f"    Carrier: {orders[0].get('Carrier Name', 'N/A')}")




