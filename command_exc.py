import subprocess
import pyautogui
import time


def switch_example(input_line):
    """
    Executes the corresponding script based on the input command.
    """
    if input_line == "execute get_profile_instant":
        exec(open('guru\\profile_instant.py').read())
        return "Option 1 selected"
    
    elif input_line == "execute get_nameplate":
        exec(open('guru\\nameplate.py').read())
        return "Option 2 selected"
    
    elif input_line == "execute get_voltage":
        exec(open('guru\\voltage.py').read())
        return "Option 3 selected"
    
    elif input_line == "execute get_net_recharge_amount":
        exec(open('guru\\total_amt_last_recharge.py').read())
        return "Option 4 selected"
    
    elif input_line == "execute get_block_load_profile_interval":
        exec(open('guru\\block_load_profile_interval.py').read())
        return "Option 5 selected"
    
    elif input_line == "execute get_MD_kW":
        exec(open('guru\\MD_kW.py').read())
        return "Option 6 selected"
    
    elif input_line == "execute get_metering_mode":
        exec(open('guru\\metering_mode.py').read())
        return "Option 7 selected"
    
    elif input_line == "execute get_payment_mode":
        exec(open('guru\\payment_mode.py').read())
        return "Option 8 selected"
    
    elif input_line == "execute get_last_token_recharge_amount":
        exec(open('guru\\last_token_recharge_amt.py').read())
        return "Option 9 selected"
    
    elif input_line == "execute get_last_token_recharge_amount_time":
        exec(open('guru\\last_token_recharge_amount_time.py').read(), globals())
        return "Option 10 selected"
    
    elif input_line == "execute get_current_balance_amount":
        exec(open('guru\\cur_balance_amount.py').read())
        return "Option 11 selected"
    
    elif input_line == "execute get_current_balance_amount_time":
        exec(open('guru\\cur_balance_time.py').read(), globals())
        return "Option 12 selected"
    
    elif input_line == "execute get_cumulative_tamper_count":
        exec(open('guru\\tamper_count.py').read())
        return "Option 13 selected"
    
    elif input_line == "execute get_live_version":
        exec(open('guru\\firmware_version.py').read())
        return "Option 14 selected"
    
    if input_line.startswith("get_block_load_profile_by_datetime"):
        exec(open('guru\\block_load.py').read())
        return "Option 15 selected"
    
    if input_line.startswith("get_daily_load_profile_by_datetime"):
        exec(open('guru\\daily_load.py').read())
        return "Option 16 selected"
    
    else:
        return "Invalid option"


# Path to your executable
path = r"C:\\Users\\krishankanty\\Desktop\\anshul\\pub_sub_emb_for_up.exe"

# Open the terminal and run the executable
command = f'start cmd /k "{path}"'
subprocess.run(command, shell=True)

# Read inputs from main.txt
with open("main.txt", "r") as file:
    inputs = file.readlines()

# Iterate through each input command
for input_line in inputs:
    # ***************************************************** Clean The txt_clear_previous_data and The pub_sub_clean_data *************************************************
    exec(open('guru\\txt_clear_previous_data.py').read())
    exec(open('pub_sub_clean_data.py').read())
    # ************************************************************* End ... *************************************************************

    # Strip newline characters and whitespace from the input line
    input_line = input_line.strip()

    # Adjust delay to allow the terminal to load
    time.sleep(2)
    pyautogui.press("enter")
    

    # Type the main command into the terminal
    pyautogui.write(input_line)
    print(f"Executing command: {input_line}")
    pyautogui.press("enter")

    # Wait for a response from the terminal (adjust timing as necessary)
    time.sleep(5)

    

    # Execute associated script based on the input command
    switch_response = switch_example(input_line)
    print(f"Switch response: {switch_response}")

# Inform the user
print("All commands are execute ...")

# Send Ctrl+C to stop any running process in the terminal
pyautogui.hotkey("ctrl", "c")
time.sleep(3)  # Small delay for Ctrl+C to take effect

# Type 'exit' to close the Command Prompt
pyautogui.write("exit")
time.sleep(2)

pyautogui.press("enter")
time.sleep(1)  # Wait before processing the next input

# Wait for 2 seconds
time.sleep(2)

# Execute the final script
exec(open('C:\\Users\\krishankanty\\Desktop\\anshul\\guru\\final_excel.py').read())

time.sleep(5)

# Execute the delete script
# exec(open('C:\\Users\\krishankanty\\Desktop\\anshul\\delete_all_xlsx_file.py').read())



print(" Task Completed ... ")


