import pyautogui
import time

print("Move the mouse to the desired location, and the coordinates will be displayed:")

# Loop to continuously print mouse position every second
try:
    while True:
        x, y = pyautogui.position()  # Get the current mouse position
        print(f"Mouse Position: X={x}, Y={y}")
        time.sleep(1)  # Wait for 1 second before getting the next position
except KeyboardInterrupt:
    print("\nProgram stopped.")
