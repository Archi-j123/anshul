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
        'active_power_W': re.search(r'active_power_W:\s*(\d+)', line),
        'MD_VA': re.search(r'MD_VA:\s*(\d+)', line),
        'frequency': re.search(r'frequency:\s*([\d.]+)', line),
        'voltage': re.search(r'voltage:\s*([\d.]+)', line),
        'neutral_current': re.search(r'neutral_current:\s*([\d.]+)', line),
        'meter_current_datetime': re.search(r'meter_current_datetime:\s*datetime.datetime\(([^)]+)\)', line),
        'load_limit_value': re.search(r'load_limit_value:\s*(\d+)', line),
        'MD_VA_datetime': re.search(r'IMPORT_MD_VA_datetime:\s*datetime.datetime\(([^)]+)\)', line),
        'cumm_tamper_count': re.search(r'cumm_tamper_count:\s*(\d+)', line),
        'load_limit_func_status': re.search(r'load_limit_func_status:\s*(\d+)', line),
        'import_Wh': re.search(r'import_Wh:\s*(\d+)', line),
        'PF': re.search(r'PF:\s*([\d.]+)', line),
        'cumm_programming_count': re.search(r'cumm_programming_count:\s*(\d+)', line),
        'apparent_power_VA': re.search(r'apparent_power_VA:\s*(\d+)', line),
        'data_type': re.search(r'data_type:\s*([A-Za-z0-9]+)', line),
        'MD_W_datetime': re.search(r'EXPORT_MD_W_datetime:\s*datetime.datetime\(([^)]+)\)', line),
        'import_VAh': re.search(r'import_VAh:\s*(\d+)', line),
        'cumm_billing_count': re.search(r'cumm_billing_count:\s*(\d+)', line),
        'export_Wh': re.search(r'export_Wh:\s*(\d+)', line),
        'export_VAh': re.search(r'export_VAh:\s*(\d+)', line),
        'MD_W': re.search(r'MD_W:\s*(\d+)', line),
        'phase_current': re.search(r'phase_current:\s*([\d.]+)', line),
        'cumm_power_on_dur_minute': re.search(r'cumm_power_on_dur_minute:\s*(\d+)', line),
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
