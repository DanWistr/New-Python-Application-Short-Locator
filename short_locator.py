import ctypes
import os
import sys
import customtkinter as ctk
import time
import threading
import subprocess
from PIL import Image
import psutil
import win32gui
import tkinter.messagebox as mb
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

### CONFIGURATIONS ###
EXE_PATH = r"C:\ShortCam II\ShortCam II.exe"
WATCH_DIR  = r"C:\ShortCam II\Record"
APP_NAME = "ShortCam II"
Camera_ID = "USB\\VID_0525&PID_A4A2\\SHORCAMII:2022" # Device Instance Path
server = r'10.45.41.79,1433\\WPH-CRE100SVR\\SQLEXPRESS' #r'DESKTOP-OVSME7R\SQLEXPRESS01' 
database = 'Short_Locator_DB' #'ShortLocator'
username = 'CREWISTRON'#'CBE100'
password = 'cre100' #'CBE100'

class JpegCreatedHandler(FileSystemEventHandler):
    def __init__(self, tk_root):
        super().__init__()
        self.tk_root = tk_root

    def on_created(self, event):
        if not event.is_directory and event.src_path.lower().endswith('.jpg'):
            # then show the window, etc.
            self.tk_root.after(0, self.tk_root.deiconify)

def monitor_directory(tk_root):

    handler  = JpegCreatedHandler(tk_root)
    observer = Observer()
    observer.schedule(handler, WATCH_DIR, recursive=False)
    time.sleep(2)
    observer.start()
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()

def on_closing():
    serial      = serial_entry.get()
    model       = model_entry.get()
    pn          = pn_entry.get()
    part        = part_entry.get()
    description = description_text_field.get("1.0", "end").strip()
    
    if serial or model or pn or part or description:
        answer = mb.askyesno(
            title="Confirm Close",
            message="Are you sure you want to close?\nUnsaved data will be lost."
        )
        if not answer:
            # they clicked “No” → do nothing
            return
        # If they clicked 'Yes', clear fields and hide
        serial_entry.delete(0, ctk.END)
        model_entry.delete(0, ctk.END)
        pn_entry.delete(0, ctk.END)
        part_entry.delete(0, ctk.END)
        description_text_field.delete("1.0", "end")
        root.withdraw()
    else:
        serial_entry.delete(0, ctk.END)
        model_entry.delete(0, ctk.END)
        pn_entry.delete(0, ctk.END)
        part_entry.delete(0, ctk.END)
        description_text_field.delete("1.0", "end")
        root.withdraw()

def is_process_running(process_name):
    for process in psutil.process_iter(attrs=['name']):
        if process.info['name'] == process_name:
            return True
    return False

def launch_with_elevated_privileges(exe_path):
    """Launch application with elevated privileges."""
    shell32 = ctypes.windll.shell32
    # Use ShellExecute to run the executable with 'runas' (elevated privileges)
    ret = shell32.ShellExecuteW(None, "runas", exe_path, None, None, 1)
    if ret <= 32:  # If the return value is 32 or less, it indicates an error.
        raise Exception(f"Failed to launch {exe_path} with elevated privileges. Return code: {ret}")

def monitor_application(process_name):
    """Monitor if the application is running and terminate the script if it closes."""
    try:
        while True:
            # Check if the process is running
            if not is_process_running(process_name):
                print(f"{process_name} has been closed. Terminating script.")
                os._exit(0)  # Forcefully terminate the script
            time.sleep(2)  # Poll every 2 seconds to reduce CPU usage
    except Exception as e:
        print(f"Error occurred during application monitoring: {e}")
        os._exit(1)  # Terminate the script in case of an error


# Set up the GUI
root = ctk.CTk()
root.title("Capture Success")
root.withdraw()

# Launch ShortCam II only if it's not already running
if not is_process_running("ShortCam II.exe"):
    try:
        launch_with_elevated_privileges(EXE_PATH)
    except Exception as e:
        print(f"Failed to launch ShortCam II: {e}")
        sys.exit(1)  # Exit the script if the application could not be launched
else:
    print("ShortCam II is already running.")

# Start a thread to monitor the application for closures
app_monitor_thread = threading.Thread(
    target=monitor_application, 
    args=("ShortCam II.exe",), 
    daemon=True  # Allows the thread to terminate when the main program exits
)
app_monitor_thread.start()

######################################################## USER INTERFACE ######################################################


# Desired window size
window_width = 1050
window_height = 800

# Center the window on screen
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
center_x = int((screen_width / 2) - (window_width / 2))
center_y = int((screen_height / 2) - (window_height / 2))
root.geometry(f"{window_width}x{window_height}+{center_x}+{center_y}")
root.resizable(False, False)

# --- Configure the main grid layout ---
root.grid_columnconfigure(0, weight=1)
root.grid_rowconfigure(5, weight=1)
root.grid_rowconfigure(9, weight=1)

# --- Header Frame (for LOGO and SEE LOGS button) ---
header_frame = ctk.CTkFrame(root, fg_color="transparent")
header_frame.grid(row=0, column=0, padx=15, pady=(20, 10), sticky="ew")
header_frame.grid_columnconfigure(0, weight=1)
header_frame.grid_columnconfigure(1, weight=0)

# Load the logo.png image
logo_image = ctk.CTkImage(Image.open("resources/SL logo.png"), size=(726, 70))
logo_label = ctk.CTkLabel(header_frame, image=logo_image, text="") 
logo_label.grid(row=0, column=0, sticky="w")

# --- SEE LOGS Button ---
see_logs_button = ctk.CTkButton(header_frame, text="SEE LOGS", width=120, height=40, font=ctk.CTkFont(size=14), fg_color="#0f6b47", hover_color="#073825")
see_logs_button.grid(row=0, column=1, sticky="e")

# --- Content Frame (for the main form fields) ---
content_frame = ctk.CTkFrame(root, fg_color="transparent")
content_frame.grid(row=1, column=0, rowspan=8, padx=50, pady=20, sticky="nsew")
content_frame.grid_columnconfigure((0, 1), weight=1)

# --- S/N and MODEL ---
serial_label = ctk.CTkLabel(content_frame, text="S/N:", font=ctk.CTkFont(size=16, weight="bold"))
serial_label.grid(row=0, column=0, sticky="w")

# Using CTkEntry for single-line input
serial_entry = ctk.CTkEntry(content_frame, height=40, font=ctk.CTkFont(size=14))
serial_entry.grid(row=1, column=0, sticky="ew", pady=5, padx=(0, 20))

model_label = ctk.CTkLabel(content_frame, text="MODEL:", font=ctk.CTkFont(size=16, weight="bold"))
model_label.grid(row=0, column=1, sticky="w", padx=(20, 0))

# Using CTkEntry for single-line input
model_entry = ctk.CTkEntry(content_frame, height=40, font=ctk.CTkFont(size=14))
model_entry.grid(row=1, column=1, sticky="ew", pady=5, padx=(20, 0))

# --- OVERHEATING COMPONENT DETAILS ---
overheating_label = ctk.CTkLabel(content_frame, text="OVERHEATING COMPONENT DETAILS", font=ctk.CTkFont(size=18, weight="bold"))
overheating_label.grid(row=2, column=0, columnspan=2, sticky="w", pady=(10, 10))

pn_label = ctk.CTkLabel(content_frame, text="PN of Detected Component:", font=ctk.CTkFont(size=14))
pn_label.grid(row=3, column=0, columnspan=2, sticky="w", padx=300, pady=(0, 5))

pn_entry = ctk.CTkEntry(content_frame, height=40, font=ctk.CTkFont(size=14))
pn_entry.grid(row=4, column=0, columnspan=2, sticky="ew", padx=300)

part_label = ctk.CTkLabel(content_frame, text="Part Type of Detected Component:", font=ctk.CTkFont(size=14))
part_label.grid(row=5, column=0, columnspan=2, sticky="w", padx=300, pady=(20, 5))

part_entry = ctk.CTkEntry(content_frame, height=40, font=ctk.CTkFont(size=14))
part_entry.grid(row=6, column=0, columnspan=2, sticky="ew", padx=300)

# --- FAILURE DESCRIPTION ---
description_label = ctk.CTkLabel(content_frame, text="FAILURE DESCRIPTION:", font=ctk.CTkFont(size=16))
description_label.grid(row=7, column=0, columnspan=2, sticky="w", pady=(20, 5))

# Textbox is correct for multi-line descriptions
description_text_field = ctk.CTkTextbox(content_frame, height=160, border_color="black", border_width=1)
description_text_field.grid(row=8, column=0, columnspan=2, sticky="ew")

# --- Buttons Frame (to group SAVE and CANCEL) ---
button_frame = ctk.CTkFrame(root, fg_color="transparent")
button_frame.grid(row=8, column=0, pady=(0, 10))  # Align frame to the left

# Variable to store radio button selection
result_var = ctk.StringVar(value="")  # You can set "PASS" or "FAIL" as default if needed

# --- Sub-frame for radio buttons aligned to left ---
radio_frame = ctk.CTkFrame(content_frame, fg_color="transparent")
radio_frame.grid(row=9, column=0, columnspan=2, sticky="w", padx=20, pady=(30, 0))  # Align to left within the frame

pass_rdbutton = ctk.CTkRadioButton(radio_frame, text="PASS", variable=result_var, value="PASS")
pass_rdbutton.grid(row=0, column=0, padx=(0, 20), pady=(0, 20))

fail_rdbutton = ctk.CTkRadioButton(radio_frame, text="FAIL", variable=result_var, value="FAIL")
fail_rdbutton.grid(row=0, column=1, pady=(0, 20))

# --- Buttons ---
save_button = ctk.CTkButton(button_frame, text="SAVE", width=140, height=40, text_color="white", font=ctk.CTkFont(size=20, weight="bold"))
save_button.grid(row=1, column=0, padx=10, sticky = "ew")

cancel_button = ctk.CTkButton(button_frame, text="CANCEL", width=140, height=40, text_color="white", fg_color="#873535", hover_color="#581A1A", font=ctk.CTkFont(size=20, weight="bold"), command=on_closing)
cancel_button.grid(row=1, column=1, padx=10, sticky = "ew")

# — start watching in the background —
monitor_thread = threading.Thread(target=monitor_directory, args=(root,), daemon=True)
monitor_thread.start()

root.protocol("WM_DELETE_WINDOW", on_closing)
root.mainloop()