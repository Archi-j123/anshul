import re
import pandas as pd
from datetime import datetime

# Step 1: Import the log file and read it
log_file_path = 'pub_sub.log'

# Read file content
with open(log_file_path, 'r', encoding='utf-8') as file:
    log_data = file.readlines()

# Initialize variables to keep track of the last command and extracted data
last_command = None
extracted_data = []

for line in log_data:
    # Check if the line contains a command starting with "get_"
    command_match = re.search(r'\b(get_\w+)', line)
    if command_match:
        # Store the command found in the current line
        last_command = command_match.group(1)
        continue  # Move to the next line to find data

    # Extract date, time, IP, and specific data fields from data rows
    date_time_match = re.search(r'\(\(\((\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\.\d{6})', line)
    ip_address_match = re.search(r'(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})', line)
    data_not_found_match = re.search(r'data not found', line)

    # Extract structured fields from data rows
    fields = {
        'cur_balance_amount': re.search(r'cur_balance_amount:\s*(\d+)', line),
        'data_type': re.search(r'data_type:\s*([A-Za-z0-9]+)', line),
        
    }

    # Collect matched data into a dictionary
    row_data = {
        'date_time': date_time_match.group(1) if date_time_match else None,
        'meter_ip_address': ip_address_match.group(1) if ip_address_match else None,
        'command_name': last_command,
        'data_not_found': bool(data_not_found_match)  # True if "data not found" was present
    }

    # Add each extracted field to row_data
    for key, match in fields.items():
        if match:
            # Convert to float if it's a numeric value
            try:
                row_data[key] = float(match.group(1))
            except ValueError:
                row_data[key] = match.group(1)

    

    # Append row_data only if it contains relevant data and has a date_time
    if row_data['date_time'] and any(value for key, value in row_data.items() if key not in ['date_time', 'meter_ip_address', 'command_name']):
        extracted_data.append(row_data)

# Step 4: Convert extracted data to a DataFrame and save it as an Excel file
df = pd.DataFrame(extracted_data)
excel_file_path = 'data.xlsx'
df.to_excel(excel_file_path, index=False)

print(f"Data successfully extracted and saved to {excel_file_path}")
