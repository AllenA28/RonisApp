import csv
from collections import defaultdict

def parse_orders(file_path):
    # Initialize a dictionary to store unique orders with categories (meat, cheese, etc.)
    orders = defaultdict(lambda: defaultdict(lambda: {'meat': set(), 'cheese': set(), 'sides': set(), 'toppings': set(), 'sauce': set()}))
    
    # Open the CSV file with error handling for encoding issues
    with open(file_path, 'r', encoding='utf-8', errors='ignore') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            # Check if 'Sent Date' exists and is not empty
            if 'Sent Date' not in row or not row['Sent Date']:
                print(f"Skipping row due to missing 'Sent Date': {row}")
                continue  # Skip this row if 'Sent Date' is missing
            
            try:
                order_id = row['Order ID']
                date = row['Sent Date'].split()[0]  # Getting only the date part (YYYY-MM-DD)
                item = row['Modifier']
                option_group = row['Option Group Name']
                
                # Skip entries with "no" or "choose" in the Modifier field
                if item.lower().startswith('no') or item.lower().startswith('choose'):
                    continue
                
                # Organize items based on their Option Group Name
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
            except IndexError as e:
                print(f"Error processing order: {row} - {e}")

    # Prepare final data structure for unique orders
    final_orders = []
    for order_id, dates in orders.items():
        for date, categories in dates.items():
            final_orders.append({
                'Order Number': order_id,
                'Date': date,
                'Meat': ', '.join(categories['meat']) if categories['meat'] else 'NA',
                'Cheese': ', '.join(categories['cheese']) if categories['cheese'] else 'NA',
                'Sides': ', '.join(categories['sides']) if categories['sides'] else 'NA',
                'Toppings': ', '.join(categories['toppings']) if categories['toppings'] else 'NA',
                'Sauce': ', '.join(categories['sauce']) if categories['sauce'] else 'NA',
            })
    
    return final_orders

# Function to write the unique orders to a CSV file
def write_orders_to_csv(orders, output_file):
    with open(output_file, 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = ['Order Number', 'Date', 'Meat', 'Cheese', 'Sides', 'Toppings', 'Sauce']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()  # Write the header
        for order in orders:
            writer.writerow(order)

# Main function to create a unique CSV from an order file
def create_unique_csv_from_order(file):
    data = parse_orders(file)
    month = file.split('_')[0]  # Extract month from filename
    output_file = f'{month}_unique_orders.csv'
    write_orders_to_csv(data, output_file)
    print(f"CSV file '{output_file}' created successfully with {len(data)} unique orders.")

# Example usage
file = 'august_2024.csv'
create_unique_csv_from_order(file)
