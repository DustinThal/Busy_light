import asyncio
import tkinter as tk
from tkinter import colorchooser, messagebox
from PIL import Image, ImageDraw
import logging
from bleak import BleakClient, BleakScanner
import winreg
import os
import json
import threading
import time
import sys
import pkgutil
import win32api
import win32con
import win32gui

# Constants
SETTINGS_FILE = "settings.json"
MIC_USAGE_KEYS = [
    r"Software\\Microsoft\\Windows\\CurrentVersion\\CapabilityAccessManager\\ConsentStore\\microphone",
    r"Software\\Microsoft\\Windows\\CurrentVersion\\CapabilityAccessManager\\ConsentStore\\microphone\\NonPackaged",
]

# Logging setup
DEBUG_MODE = True  # Set to False to disable logging
if DEBUG_MODE:
    logging.basicConfig(filename="app.log", level=logging.DEBUG, format="%(asctime)s - %(levelname)s - %(message)s")
else:
    logging.basicConfig(level=logging.CRITICAL)  # Suppress all logs if DEBUG_MODE is False

# Global variables
mic_in_use = False
bluetooth_connected = False
client = None
mic_color = "255,0,0"  # Default red for mic in use
idle_color = "0,255,0"  # Default green for mic idle
last_device_address = None
running = False
lock = threading.Lock()  # Synchronization lock
last_color_sent = None  # Track the last color sent
time_last_sent = 0  # Track time since the last RGB data was sent
last_mic_status = None  # Track the last known microphone status
tray_icon = None
main_loop = None
bluetooth_filter = "busy_light_"  # Default filter for Bluetooth devices

# Windows tray-specific variables
TRAY_ICON_ID = 1
WND_CLASS_NAME = "BusyLightControllerTray"

tray_hwnd = None
tray_icon_data = None
hicon = None

def initialize_event_loop():
    global main_loop
    if main_loop is None:
        main_loop = asyncio.new_event_loop()
        threading.Thread(target=main_loop.run_forever, daemon=True).start()

initialize_event_loop()

def log_environment_info():
    logging.info(f"Python version: {sys.version}")
    logging.info(f"Loaded modules: {[module.name for module in pkgutil.iter_modules()]}")

def save_settings():
    settings = {"mic_color": mic_color, "idle_color": idle_color, "bluetooth_filter": bluetooth_filter}
    with open(SETTINGS_FILE, "w") as f:
        json.dump(settings, f)

def load_settings():
    global mic_color, idle_color, bluetooth_filter
    if os.path.exists(SETTINGS_FILE):
        with open(SETTINGS_FILE, "r") as f:
            settings = json.load(f)
            mic_color = settings.get("mic_color", mic_color)
            idle_color = settings.get("idle_color", idle_color)
            bluetooth_filter = settings.get("bluetooth_filter", bluetooth_filter)

def create_icon():
    """Create the icon image."""
    width, height = 64, 64
    image = Image.new("RGB", (width, height), "blue")
    draw = ImageDraw.Draw(image)
    draw.ellipse((16, 16, 48, 48), fill="red")
    icon_path = "tray_icon.ico"
    image.save(icon_path)
    return win32gui.LoadImage(None, icon_path, win32con.IMAGE_ICON, 64, 64, win32con.LR_LOADFROMFILE)

def on_tray_event(hwnd, msg, wparam, lparam):
    """Handle tray icon events."""
    if lparam == win32con.WM_LBUTTONUP:  # Left-click
        show_window()
    elif lparam == win32con.WM_RBUTTONUP:  # Right-click
        show_tray_menu(hwnd)
    return win32gui.DefWindowProc(hwnd, msg, wparam, lparam)

def create_tray_icon():
    """Create the system tray icon."""
    global tray_hwnd, tray_icon_data, hicon

    if tray_hwnd:  # If the tray icon already exists, don't recreate it
        return

    # Ensure class is registered only once
    try:
        wc = win32gui.WNDCLASS()
        wc.hInstance = win32api.GetModuleHandle(None)
        wc.lpszClassName = WND_CLASS_NAME
        wc.lpfnWndProc = on_tray_event
        win32gui.RegisterClass(wc)
    except Exception as e:
        logging.debug(f"Class registration error (probably already registered): {e}")

    # Create window
    tray_hwnd = win32gui.CreateWindow(WND_CLASS_NAME, WND_CLASS_NAME, 0, 0, 0, 0, 0, 0, 0, None, None)

    # Create and add the tray icon
    hicon = create_icon()
    tray_icon_data = (
        tray_hwnd,
        TRAY_ICON_ID,
        win32gui.NIF_ICON | win32gui.NIF_MESSAGE | win32gui.NIF_TIP,
        win32con.WM_USER + 20,
        hicon,
        "Busy Light Controller",
    )
    win32gui.Shell_NotifyIcon(win32gui.NIM_ADD, tray_icon_data)

def show_tray_menu(hwnd):
    """Show the right-click menu."""
    menu = win32gui.CreatePopupMenu()
    win32gui.AppendMenu(menu, win32con.MF_STRING, 1, "Show Window")
    win32gui.AppendMenu(menu, win32con.MF_STRING, 2, "Exit")

    # Get cursor position and show the menu
    x, y = win32gui.GetCursorPos()
    command = win32gui.TrackPopupMenu(menu, win32con.TPM_RETURNCMD | win32con.TPM_NONOTIFY, x, y, 0, hwnd, None)

    if command == 1:  # Show Window
        show_window()
    elif command == 2:  # Exit
        on_exit()

def show_window():
    """Show the main application window."""
    global window
    window.deiconify()
    window.lift()
    window.focus_force()

def on_exit():
    """Handle application exit."""
    global tray_hwnd, tray_icon_data, hicon

    # Remove the tray icon
    if tray_icon_data:
        win32gui.Shell_NotifyIcon(win32gui.NIM_DELETE, tray_icon_data)
        tray_icon_data = None

    # Destroy the window
    if tray_hwnd:
        win32gui.DestroyWindow(tray_hwnd)
        tray_hwnd = None

    # Destroy the icon
    if hicon:
        win32gui.DestroyIcon(hicon)
        hicon = None

    # Exit the program
    sys.exit(0)

def check_microphone_usage():
    global mic_in_use, running, last_mic_status
    first_run = True
    while running:
        found_in_use = False
        for root_key in MIC_USAGE_KEYS:
            try:
                reg_key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, root_key)
                i = 0
                while True:
                    try:
                        subkey_name = winreg.EnumKey(reg_key, i)
                        subkey_path = f"{root_key}\\{subkey_name}"
                        subkey = winreg.OpenKey(winreg.HKEY_CURRENT_USER, subkey_path)
                        last_used_time_stop, _ = winreg.QueryValueEx(subkey, "LastUsedTimeStop")
                        if last_used_time_stop == 0:
                            found_in_use = True
                            break
                    except FileNotFoundError:
                        pass
                    except OSError:
                        break
                    i += 1
            except FileNotFoundError:
                pass
            except PermissionError:
                logging.error(f"Permission denied when accessing: {root_key}")

        with lock:
            mic_in_use = found_in_use
            if first_run or mic_in_use != last_mic_status:
                current_color = mic_color if mic_in_use else idle_color
                asyncio.run_coroutine_threadsafe(send_color(current_color), main_loop)
                last_mic_status = mic_in_use
                first_run = False
        time.sleep(3)

async def send_color(color):
    global client, bluetooth_connected, last_color_sent, time_last_sent
    current_time = time.time()

    if client and bluetooth_connected:
        if color != last_color_sent or (current_time - time_last_sent) > 180:
            try:
                await client.write_gatt_char("abcd1234-abcd-1234-abcd-12345678abcd", color.encode())
                logging.debug(f"Sent color: {color}")
                last_color_sent = color
                time_last_sent = current_time
            except Exception as e:
                logging.error(f"Error sending color: {e}")

def update_status():
    global last_mic_status

    with lock:
        current_color = mic_color if mic_in_use else idle_color
        if mic_in_use != last_mic_status:
            mic_status_label.config(text=f"Microphone Status: {'In Use' if mic_in_use else 'Idle'}")
            asyncio.run_coroutine_threadsafe(send_color(current_color), main_loop)
            last_mic_status = mic_in_use

    bt_status_label.config(text=f"Bluetooth Status: {'Connected' if bluetooth_connected else 'Disconnected'}")
    mic_status_label.config(text=f"Microphone Status: {'In Use' if mic_in_use else 'Idle'}")

    if bluetooth_connected:
        bluetooth_button.config(state=tk.DISABLED)
        disconnect_button.config(state=tk.NORMAL)
    else:
        bluetooth_button.config(state=tk.NORMAL)
        disconnect_button.config(state=tk.DISABLED)

    window.after(1000, update_status)

def pick_color(use_mic):
    global mic_color, idle_color
    color_code = colorchooser.askcolor(title="Choose color")[0]
    if color_code:
        color = f"{int(color_code[0])},{int(color_code[1])},{int(color_code[2])}"
        if use_mic:
            mic_color = color
        else:
            idle_color = color
        save_settings()

def on_close():
    global tray_hwnd, tray_icon_created
    window.withdraw()  # Minimize to tray

    if not tray_hwnd:  # Create the tray icon only if it doesn't exist
        create_tray_icon()

def on_minimize(event):
    on_close()
    
async def find_device():
    try:
        devices = await BleakScanner.discover()
        for device in devices:
            if device.name and bluetooth_filter in device.name:
                return device.address
    except Exception as e:
        logging.error(f"Error during Bluetooth discovery: {e}")
    return None

async def connect_device():
    global client, bluetooth_connected
    mac_address = await find_device()
    if mac_address:
        try:
            client = BleakClient(mac_address)
            await client.connect()
            bluetooth_connected = client.is_connected
            if bluetooth_connected:
                bt_status_label.config(text="Bluetooth Status: Connected")
                disconnect_button.config(state=tk.NORMAL)
                bluetooth_button.config(state=tk.DISABLED)
                asyncio.run_coroutine_threadsafe(monitor_session_status(), main_loop)
        except Exception as e:
            bluetooth_connected = False
            logging.error(f"Connection failed: {e}")
            handle_closing_session("CDConnection failed")
    else:
        messagebox.showwarning("Device Not Found", "No suitable device found.")

def disconnect_bluetooth_force():
    global client, bluetooth_connected
    logging.error(f"disconnect_bluetooth: started")
    try:
        asyncio.run_coroutine_threadsafe(client.disconnect(), main_loop)
        bluetooth_connected = False
    except Exception as e:
        logging.error(f"Error sending 'disconnect' command: {e}")

def disconnect_bluetooth():
    global client, bluetooth_connected
    logging.error(f"disconnect_bluetooth: started")
    if client and bluetooth_connected:
        try:
            asyncio.run_coroutine_threadsafe(client.disconnect(), main_loop)
            bluetooth_connected = False
        except Exception as e:
            logging.error(f"Error sending 'disconnect' command: {e}")

def start_microphone_identification():
    global running
    if not running:
        running = True
        threading.Thread(target=check_microphone_usage, daemon=True).start()
        start_button.config(state=tk.DISABLED)
        stop_button.config(state=tk.NORMAL)

def stop_microphone_identification():
    global running
    if running:
        running = False
        start_button.config(state=tk.NORMAL)
        stop_button.config(state=tk.DISABLED)

# Handle session closure
def handle_closing_session(reason):
    global bluetooth_connected
    if not bluetooth_connected:  # Prevent redundant actions
        return
    bluetooth_connected = False
    logging.warning(f"Session closed: {reason}")
    bt_status_label.config(text=f"Bluetooth Status: Disconnected ({reason})")
    disconnect_button.config(state=tk.DISABLED)
    bluetooth_button.config(state=tk.NORMAL)
    disconnect_bluetooth_force()

async def monitor_session_status():
    global bluetooth_connected
    while bluetooth_connected:
        try:
            if client and not client.is_connected:
                logging.warning("Connection lost during monitoring.")
                handle_closing_session("MSSConnection lost (monitor)")
                break
        except Exception as e:
            logging.error(f"Error during status monitoring: {e}")
        await asyncio.sleep(2)

def main():
    load_settings()
    global window, bluetooth_button, disconnect_button, mic_status_label, bt_status_label, start_button, stop_button, mic_color_button, idle_color_button

    window = tk.Tk()
    window.title("Busy Light Controller")
    window.protocol("WM_DELETE_WINDOW", lambda: on_exit())

    # Bind minimize button
    window.bind("<Unmap>", lambda e: on_close() if window.state() == "iconic" else None)

    bt_status_label = tk.Label(window, text="Bluetooth Status: Disconnected", font=("Arial", 14))
    bt_status_label.pack(pady=10)

    bluetooth_frame = tk.Frame(window)
    bluetooth_frame.pack(pady=5)

    bluetooth_button = tk.Button(bluetooth_frame, text="Connect to Bluetooth", command=lambda: asyncio.run_coroutine_threadsafe(connect_device(), main_loop))
    bluetooth_button.pack(side=tk.LEFT, padx=5)

    disconnect_button = tk.Button(bluetooth_frame, text="Disconnect Bluetooth", command=disconnect_bluetooth, state=tk.NORMAL)
    disconnect_button.pack(side=tk.LEFT, padx=5)

    mic_status_label = tk.Label(window, text="Microphone Status: Unknown", font=("Arial", 14))
    mic_status_label.pack(pady=10)

    start_button = tk.Button(window, text="Start Mic Identification", command=start_microphone_identification)
    start_button.pack(pady=5)

    stop_button = tk.Button(window, text="Stop Mic Identification", command=stop_microphone_identification, state=tk.DISABLED)
    stop_button.pack(pady=5)

    color_frame = tk.Frame(window)
    color_frame.pack(pady=10)

    mic_color_button = tk.Button(color_frame, text="Set Mic In Use Color", command=lambda: pick_color(True))
    mic_color_button.pack(side=tk.LEFT, padx=5)

    idle_color_button = tk.Button(color_frame, text="Set Mic Idle Color", command=lambda: pick_color(False))
    idle_color_button.pack(side=tk.LEFT, padx=5)

    bluetooth_filter_frame = tk.Frame(window)
    bluetooth_filter_frame.pack(pady=10)

    tk.Label(bluetooth_filter_frame, text="Bluetooth Filter:", font=("Arial", 12)).pack(side=tk.LEFT, padx=5)
    bluetooth_filter_entry = tk.Entry(bluetooth_filter_frame)
    bluetooth_filter_entry.insert(0, bluetooth_filter)
    bluetooth_filter_entry.pack(side=tk.LEFT, padx=5)

    def update_filter():
        global bluetooth_filter
        bluetooth_filter = bluetooth_filter_entry.get()
        save_settings()

    filter_button = tk.Button(bluetooth_filter_frame, text="Update Filter", command=update_filter)
    filter_button.pack(side=tk.LEFT, padx=5)

    update_status()
    window.mainloop()

if __name__ == "__main__":
    try:
        create_tray_icon()
        main()
    except Exception as e:
        logging.error(f"Application error: {e}")
        messagebox.showerror("Application Error", f"An error occurred: {e}")
    finally:
        on_exit()
