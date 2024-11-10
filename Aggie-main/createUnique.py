#this file has a function at the botom that takes in a csv file and writes a new csv file with all the unique orders


import csv
from collections import defaultdict

def parse_orders(file_path):
    orders = defaultdict(lambda: defaultdict(lambda: {'meat': set(), 'cheese': set(), 'sides': set(), 'toppings': set(), 'sauce': set()}))
    
    with open(file_path, 'r') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            order_id = row['Order ID']
            date = row['Sent Date'].split()[0]  # Getting the date only
            item = row['Modifier']
            option_group = row['Option Group Name']
            
            # Filter out unwanted entries
            if item.lower().startswith('no') or item.lower().startswith('choose'):
                continue
            
            # Organizing the items based on the option group
            if 'Meats' in option_group:
                orders[order_id][date]['meat'].add(item)
            elif 'Cheese' in option_group:
                orders[order_id][date]['cheese'].add(item)
            elif 'Side' in option_group:
                orders[order_id][date]['sides'].add(item)
            elif 'Toppings' in option_group:
                orders[order_id][date]['toppings'].add(item)
            elif 'Drizzles' in option_group:
                orders[order_id][date]['sauce'].add(item)

    # Prepare final data structure
    final_orders = []
    for order_id, dates in orders.items():
        for date, categories in dates.items():
            final_orders.append({
                'order number': order_id,
                'date': date,
                'meat': ', '.join(categories['meat']) if categories['meat'] else 'NA',
                'cheese': ', '.join(categories['cheese']) if categories['cheese'] else 'NA',
                'sides': ', '.join(categories['sides']) if categories['sides'] else 'NA',
                'toppings': ', '.join(categories['toppings']) if categories['toppings'] else 'NA',
                'sauce': ', '.join(categories['sauce']) if categories['sauce'] else 'NA',
            })
    
    return final_orders

def write_orders_to_csv(orders, output_file):
    with open(output_file, 'w', newline='') as csvfile:
        fieldnames = ['order number', 'date', 'meat', 'cheese', 'sides', 'toppings', 'sauce']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        
        writer.writeheader()  # Write the header
        for order in orders:
            writer.writerow(order)

# Parse the orders
def create_unique_csv_from_order(file):
    data = parse_orders(file)
    month = file[:file.index('_')]
    if file.index('_') == -1:
        month = 'unknownMonth'
    output = month + '_unique_orders.csv'
    write_orders_to_csv(data, output)


file = 'june_2024.csv'
numbers = parse_orders(file)
# Write to new CSV file
output_file_path = 'july unique_orders.csv'
create_unique_csv_from_order(file)

print(f"CSV file '{output_file_path}' created successfully with {len(numbers)} unique orders.")