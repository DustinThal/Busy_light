import winreg
import time
import serial
import serial.tools.list_ports
import threading
import tkinter as tk
from tkinter import messagebox
import logging
import pystray
from pystray import MenuItem as item
from PIL import Image, ImageDraw

# Set up logging
DEBUG_MODE = False  # Set this to False to disable debugging (logs and console messages)
logging.basicConfig(filename='log.txt', level=logging.DEBUG if DEBUG_MODE else logging.INFO, format='%(asctime)s - %(message)s')

# Global variables
esp_connected = None
microphone_in_use = False
esp_port = None
serial_connection = None  # To store the serial connection object
running = False  # To track if the microphone identification is running
SHOW_ARDUINO_RESPONSE = False  # Set this to False to hide Arduino responses in the GUI
tray_icon = None  # For system tray icon
microphone_thread = None  # To keep track of the microphone-checking thread
last_color_sent = None  # To prevent redundant command sending

# Function to create the system tray icon
def create_tray_icon():
    icon_image = Image.new('RGB', (64, 64), color=(255, 255, 255))
    draw = ImageDraw.Draw(icon_image)
    draw.rectangle([(0, 0), (64, 64)], fill="blue")  # Draw a simple blue square as the icon

    tray_icon = pystray.Icon("Busylight", icon_image, menu=pystray.Menu(
        item('Restore', restore_window),
        item('Quit', quit_program)
    ))
    tray_icon.run()

def restore_window(icon, item):
    """Restore the Tkinter window when clicked in the tray."""
    window.deiconify()  # Show the Tkinter window again
    window.update_idletasks()
    logging.debug("Window restored.")

def quit_program(icon, item):
    """Gracefully quit the application."""
    icon.stop()
    logging.debug("System Tray Icon stopped.")

    # Gracefully exit Tkinter and the application
    window.quit()
    window.destroy()

    # Stop background threads
    global running
    running = False
    if microphone_thread:
        microphone_thread.join()  # Ensure the background thread is stopped

def check_microphone_usage():
    global microphone_in_use
    root_key = r"Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\microphone"
    try:
        reg_key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, root_key)
        logging.debug("Checking microphone usage from the registry.")
    except FileNotFoundError:
        logging.error("Registry key not found. Unable to check microphone usage.")
        return

    while running:
        microphone_in_use = False
        try:
            i = 0
            while True:
                subkey_name = winreg.EnumKey(reg_key, i)
                subkey_path = f"{root_key}\\{subkey_name}"
                try:
                    subkey = winreg.OpenKey(winreg.HKEY_CURRENT_USER, subkey_path)
                    last_used_time_stop, _ = winreg.QueryValueEx(subkey, "LastUsedTimeStop")
                    if last_used_time_stop == 0:
                        microphone_in_use = True
                        break
                except FileNotFoundError:
                    pass
                i += 1
        except OSError:
            pass

        # Now check the NonPackaged directory
        nonpackaged_key = r"Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\microphone\NonPackaged"
        try:
            nonpackaged_reg_key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, nonpackaged_key)
            nonpackaged_i = 0
            while True:
                subkey_name = winreg.EnumKey(nonpackaged_reg_key, nonpackaged_i)
                nonpackaged_subkey_path = f"{nonpackaged_key}\\{subkey_name}"
                try:
                    nonpackaged_subkey = winreg.OpenKey(winreg.HKEY_CURRENT_USER, nonpackaged_subkey_path)
                    last_used_time_stop, _ = winreg.QueryValueEx(nonpackaged_subkey, "LastUsedTimeStop")
                    if last_used_time_stop == 0:
                        microphone_in_use = True
                        break
                except FileNotFoundError:
                    pass
                nonpackaged_i += 1
        except OSError:
            pass

        logging.debug(f"Microphone status: {'In Use' if microphone_in_use else 'Not in Use'}")
        time.sleep(3)  # Adjusted to check every 3 seconds

def detect_esp32c6():
    global esp_port
    connected_ports = []
    for port in serial.tools.list_ports.comports():
        if 'USB' in port.description:
            connected_ports.append(port.device)

    if len(connected_ports) == 1:
        esp_port = connected_ports[0]
        logging.debug(f"Detected ESP32c6 on port: {esp_port}")
    else:
        logging.error("Multiple or no USB devices found. Please select the ESP32c6 port manually.")

def send_color_to_esp32(color):
    global serial_connection, last_color_sent
    if serial_connection is None:
        logging.error("No ESP32c6 connected.")
        return

    if color == last_color_sent:
        logging.debug("Command not sent as it matches the last sent color.")
        return

    try:
        logging.debug(f"Attempting to send command: {color}")
        serial_connection.write(color.encode())  # Send the color change ('R', 'G', 'W')
        time.sleep(0.1)  # Short delay to give the Arduino time to respond
        response = serial_connection.readline().decode('utf-8').strip()  # Read the echo response from Arduino

        if response and SHOW_ARDUINO_RESPONSE:  # Only show non-empty response if enabled
            logging.debug(f"Received from Arduino: {response}")
            response_box.insert(tk.END, f"Arduino: {response}\n")  # Display the Arduino response in the text box
            response_box.yview(tk.END)  # Scroll to the bottom
        elif not SHOW_ARDUINO_RESPONSE:
            logging.debug("Arduino response display is off.")
        else:
            logging.error("No response from Arduino")

        last_color_sent = color  # Update the last sent color
    except Exception as e:
        logging.error(f"Error sending data to ESP32c6: {e}")

def update_status():
    mic_status = "In Use" if microphone_in_use else "Not in Use"
    mic_status_label.config(text=f"Microphone Status: {mic_status}")
    if esp_port is not None:
        com_port_label.config(text=f"Connected to: {esp_port}")
    else:
        com_port_label.config(text="No COM Port Connected")

    # Send color command to ESP32 based on microphone status
    if microphone_in_use:
        send_color_to_esp32('R')  # Red for in use
    else:
        send_color_to_esp32('G')  # Green for not in use

    logging.debug(f"Status updated: Microphone is {'in use' if microphone_in_use else 'not in use'}.")
    window.after(1000, update_status)  # Update every 1 second

def create_window():
    global window, mic_status_label, com_port_label, response_box, start_button, stop_button
    window = tk.Tk()
    window.title("USB Busy Light")

    mic_status_label = tk.Label(window, text="Microphone Status: Not Monitoring", font=("Arial", 14))
    mic_status_label.pack(pady=10)

    com_port_label = tk.Label(window, text="No COM Port Connected", font=("Arial", 14))
    com_port_label.pack(pady=10)

    start_button = tk.Button(window, text="Start Microphone Identification", font=("Arial", 12), command=start_microphone_identification)
    start_button.pack(pady=10)

    stop_button = tk.Button(window, text="Stop Microphone Identification", font=("Arial", 12), command=stop_microphone_identification, state=tk.DISABLED)
    stop_button.pack(pady=10)

    response_box = tk.Text(window, height=10, width=50, font=("Arial", 12))
    response_box.pack(pady=10)
    if not SHOW_ARDUINO_RESPONSE:
        response_box.pack_forget()  # Hide the response box initially if SHOW_ARDUINO_RESPONSE is False

    # Bind minimize event to hide the window
    window.protocol("WM_DELETE_WINDOW", minimize_to_tray)

    window.withdraw()
    tray_thread = threading.Thread(target=create_tray_icon)
    tray_thread.daemon = True
    tray_thread.start()

    window.after(1000, update_status)
    window.mainloop()

def minimize_to_tray():
    """Minimize the window to the system tray."""
    window.withdraw()
    logging.debug("Window minimized to tray.")

def start_microphone_identification():
    global running, serial_connection
    running = True
    logging.debug("Starting microphone identification.")
    detect_esp32c6()

    try:
        serial_connection = serial.Serial(esp_port, 115200, timeout=1)
        logging.debug(f"Successfully opened serial port: {esp_port}")
    except Exception as e:
        logging.error(f"Error opening serial port: {e}")
        return

    global microphone_thread
    microphone_thread = threading.Thread(target=check_microphone_usage)
    microphone_thread.daemon = True
    microphone_thread.start()

    # Update button states
    start_button.config(state=tk.DISABLED)
    stop_button.config(state=tk.NORMAL)

def stop_microphone_identification():
    global running, serial_connection
    running = False
    logging.debug("Stopping microphone identification.")

    if serial_connection is not None:
        serial_connection.close()
        logging.debug(f"Serial port {esp_port} closed.")

    # Update button states
    start_button.config(state=tk.NORMAL)
    stop_button.config(state=tk.DISABLED)

if __name__ == "__main__":
    create_window()
