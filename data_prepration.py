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
    
    # Check for "data not found" status
    data_not_found_match = re.search(r'data not found', line)

    # Extract structured fields from data rows
    fields = {
        'index': re.search(r'index:\s*(\d+)', line),
        'measured_current': re.search(r'measured_current:\s*([\d.]+)', line),
        'data_type': re.search(r'data_type:\s*([A-Za-z0-9]+)', line),
        'cumm_tamper_count': re.search(r'cumm_tamper_count:\s*(\d+)', line),
        'cumm_energy_Wh': re.search(r'cumm_energy_Wh:\s*(\d+)', line),
        'PF': re.search(r'PF:\s*([\d.]+)', line),
        'total_packet': re.search(r'total_packet:\s*(\d+)', line),
        'event_code': re.search(r'event_code:\s*(\d+)', line),
        'event_datetime': re.search(r'event_datetime:\s*datetime.datetime\(([^)]+)\)', line),
        # 'voltage': re.search(r'voltage:\s*([\d.]+)', line),
        'category': re.search(r'category:\s*([A-Za-z0-9]+)', line),
        'manufacturer_name': re.search(r'manufacturer_name:\s*([^,]+)', line),
        'current_rating': re.search(r'current_rating:\s*\(([^)]+)\)', line),
        'manufacturing_year': re.search(r'manufacturing_year:\s*(\d+)', line),
        'firmware_version': re.search(r'firmware_version:\s*(\d+)', line),
        'SM_device_id': re.search(r'SM_device_id:\s*([A-Za-z0-9]+)', line),
        'type': re.search(r'type:\s*(\d+)', line),
        'Total_Programming_Count': re.search(r'Total\s+Programming\s+Count:\s*(\d+)', line),
        'First_Power_Up_Timestamp': re.search(r'First\s+Power\s+Up\s+Timestamp:\s*([\d-]+\s+[\d:]+)', line),
        'Total_Meter_Power_Up_Count': re.search(r'Total\s+Meter\s+Power-Up\s+Count:\s*(\d+)', line),
        'Last_Power_Failure_Duration': re.search(r'Last\s+Power\s+Failure\s+Duration:\s*([\d\s\w]+)', line),
        'Total_Billing_Count': re.search(r'Total\s+Billing\s+Count:\s*(\d+)', line),
        'Demand_Integration_Period': re.search(r'Demand\s+Integration\s+Period\s+\(Mins\):\s*(\d+)', line),
        'Total_Power_On_Duration': re.search(r'Total\s+Power\s+On\s+Duration:\s*([\d\s\w]+)', line),
        'Power_On_Timestamp': re.search(r'Power\s+On\s+Timestamp:\s*([\d-]+\s+[\d:]+)', line),
        'meter_clock': re.search(r'meter_clock:\s*datetime.datetime\(([^)]+)\)', line),
        'PCP_value': re.search(r'PCP_value:\s*(\d+)', line),
        'import_VAh': re.search(r'import_VAh:\s*(\d+)', line),
        'export_VAh': re.search(r'export_VAh:\s*(\d+)', line),
        'import_Wh': re.search(r'import_Wh:\s*(\d+)', line),
        'export_Wh': re.search(r'export_Wh:\s*(\d+)', line),
        'cur_balance_amount': re.search(r'cur_balance_amount:\s*(\d+)', line),
        'cur_balance_time': re.search(r'cur_balance_time:\s*\(([^)]+)\)', line),
        'MD_W': re.search(r'MD_W:\s*(\d+)', line),  # Maximum Demand in watts
        
        'block_active_energy_exp': re.search(r'block_active_energy_exp:\s*([\d.]+)', line),
        'date_time1': re.search(r'date_time1:\s*([\d-]+\s+[\d:]+)', line),
        'block_apparent_energy2': re.search(r'block_apparent_energy2:\s*([\d.]+)', line),
        'temperature': re.search(r'temperature:\s*([\d.]+)', line),
        'block_real_energy': re.search(r'block_real_energy:\s*([\d.]+)', line),
        'block_apparent_energy_exp1': re.search(r'block_apparent_energy_exp1:\s*([\d.]+)', line),
        'block_active_energy_exp1': re.search(r'block_active_energy_exp1:\s*([\d.]+)', line),
        'avg_voltage': re.search(r'avg_voltage:\s*([\d.]+)', line),
        'avg_current': re.search(r'avg_current:\s*([\d.]+)', line),
        'block_real_energy1': re.search(r'block_real_energy1:\s*([\d.]+)', line),
        'temperature1': re.search(r'temperature1:\s*([\d.]+)', line),
        'block_apparent_energy_exp': re.search(r'block_apparent_energy_exp:\s*([\d.]+)', line),
        'current1': re.search(r'current1:\s*([\d.]+)', line),
        'block_apparent_energy': re.search(r'block_apparent_energy:\s*([\d.]+)', line),

        'metering_mode': re.search(r'metering_mode:\s*(\d+)', line),
        'payment_mode': re.search(r'payment_mode:\s*(\d+)', line),

        'last_token_recharge_time': re.search(r'last_token_recharge_time:\s*\(([^)]+)\)', line),
        'total_amt_last_recharge': re.search(r'total_amt_last_recharge:\s*(\d+)', line),
        'last_token_recharge_amt': re.search(r'last_token_recharge_amt:\s*(\d+)', line),
        'metering_mode': re.search(r'metering_mode:\s*(\d+)', line),
        'payment_mode': re.search(r'payment_mode:\s*(\d+)', line),
        'net_recharge_amount': re.search(r'net_recharge_amount:\s*(\d+)', line),
        'cur_balance_time': re.search(r'cur_balance_time:\s*\(([^)]+)\)', line),
    
    }

    # Collect matched data into a dictionary
    row_data = {
        'date_time': date_time_match.group(1) if date_time_match else None,
        'meter_ip_address': ip_address_match.group(1) if ip_address_match else None,
        'command_name': last_command,
        'data_not_found': bool(data_not_found_match)  # True if "data not found" was present
    }

       # Append to the list if date_time is captured
    if row_data['date_time']:
        extracted_data.append(row_data)

    # Add each extracted field to row_data
    for key, match in fields.items():
        row_data[key] = match.group(1) if match else None

    # Convert `event_datetime` and `meter_clock` to standard datetime format if present
    if row_data['event_datetime']:
        try:
            row_data['event_datetime'] = datetime.strptime(row_data['event_datetime'], '%Y, %m, %d, %H, %M, %S')
        except ValueError:
            pass  # Handle any conversion errors if necessary
    if row_data['meter_clock']:
        try:
            row_data['meter_clock'] = datetime.strptime(row_data['meter_clock'], '%Y, %m, %d, %H, %M, %S')
        except ValueError:
            pass  # Handle any conversion errors if necessary

    # Append row_data only if it contains relevant data
    if any(value for key, value in row_data.items() if key not in ['date_time', 'meter_ip_address', 'command_name']):
        extracted_data.append(row_data)

# Step 4: Convert extracted data to a DataFrame and save it as an Excel file
df = pd.DataFrame(extracted_data)
excel_file_path = 'data.xlsx'
df.to_excel(excel_file_path, index=False)

print(f"Data successfully extracted and saved to {excel_file_path}")
