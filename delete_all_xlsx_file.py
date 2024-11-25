import os
import glob

# Define the folder path
downloads_folder = os.path.expanduser("~/Downloads")

# List of filenames to delete (without extensions)
files_to_delete = [
    "block_load_profile_interval",
    "block_load",
    "last_token_recharge_amt",
    "total_amt_last_recharge",
    "cur_balance_amount",
    "PCP_value",
    "daily_load",
    "cumm_tamper_count",
    "cur_balance_time",
    "firmware_version",
    "profile_instant",
    "MD_kW",
    "last_token_recharge_time",
    "voltage",
    "metering_mode",
    "payment_mode",
    "nameplate",
    "power_event",
    "tamper_count",
    "voltage"
]

# Iterate over files and delete if they exist
for file_name in files_to_delete:
    file_pattern = os.path.join(downloads_folder, f"{file_name}*.xlsx")
    files = glob.glob(file_pattern)  # Search for matching files
    for file_path in files:
        try:
            os.remove(file_path)
            # print(f"Deleted: {file_path}")
        except Exception as e:
            print(f"Failed to delete {file_path}: {e}")

print("Cleanup complete.")
