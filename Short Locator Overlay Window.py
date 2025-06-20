import customtkinter as ctk
import subprocess
import time
import psutil
from PIL import Image
import pygetwindow
import win32gui
import pyautogui
import pyodbc
import cv2
import threading
import os
import winreg
import sys
from datetime import datetime

from pathlib import Path


### CONFIGURATIONS ###
EXE_PATH = r"C:\ShortCam II\ShortCam II.exe"
APP_NAME = "ShortCam II"
Camera_ID = "USB\\VID_0525&PID_A4A2\\SHORCAMII:2022" # Device Instance Path
server = r'10.45.41.79,1433\\WPH-CRE100SVR\\SQLEXPRESS' #r'DESKTOP-OVSME7R\SQLEXPRESS01' 
database = 'Short_Locator_DB' #'ShortLocator'
username = 'CREWISTRON'#'CBE100'
password = 'cre100' #'CBE100'

input_sn = None
get_pn = None
get_part_type = None
row_data = None
query_result = None
date_now = None

capture_button_located = None
capture_success_located = None
capture_failed_located = None
elapsed_time = 0
timeout = 10
interval: int = 1

#####################################################
########## Auto-Py-to-Exe Preparation ###############
#####################################################

def is_odbc_driver_installed(driver_name="ODBC Driver 17 for SQL Server"):
    try:
        path = r"SOFTWARE\ODBC\ODBCINST.INI\{}".format(driver_name)
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path):
            return True
    except FileNotFoundError:
        return False

def install_odbc_driver():
    installer_path = os.path.join(os.getcwd(), "msodbcsql17.msi")
    if os.path.exists(installer_path):
        subprocess.run(["msiexec", "/i", installer_path, "/quiet", "/norestart"], check=True)
    else:
        raise FileNotFoundError("ODBC installer not found!")

# Use this early in your script
if not is_odbc_driver_installed():
    install_odbc_driver()

def resource_path(relative_path):
    """ Get the absolute path to resource, works for dev and for PyInstaller EXE """
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

#####################################################
########## ShortCam II UI Interaction Part ##########
#####################################################

# Check if ShortCam II is already running
def is_process_running(process_name):
    for process in psutil.process_iter(attrs=['name']):
        if process.info['name'] == process_name:
            return True
    return False

# Launch ShortCam II only if it's not already running
if not is_process_running("ShortCam II.exe"):
    try:
        subprocess.Popen(EXE_PATH)
    except Exception as e:
        print(f"Failed to launch ShortCam II: {e}")
else:
    print("ShortCam II is already running.")

# Find the ShortCam II window
def find_window():
    hwnd = win32gui.FindWindow(None, APP_NAME)
    return hwnd if hwnd else None
print("Looking for ShortCam II...")
ShortCam_win = None
while not ShortCam_win:
    ShortCam_win = find_window()
    time.sleep(0.5)  # Faster response time
print(f"ShortCam II found: {ShortCam_win}")

# # Enable Thermal Camera through Device Console
# def enable_thermalcam(Camera_ID):
#     try:
#         cmd = f"powershell -Command \"Enable-PnpDevice -InstanceId '{Camera_ID}' -Confirm:$false\""
#         subprocess.run(cmd, shell=True)
#         print("Enable success")
#     except subprocess.CalledProcessError as e:
#         print(f"Error occurred while disabling device: {e}")


# # Disable Thermal Camera through Device Console
# def disable_thermalcam(Camera_ID):
#     try:
#         cmd = f"powershell -Command \"Disable-PnpDevice -InstanceId '{Camera_ID}' -Confirm:$false\""
#         subprocess.run(cmd, shell=True)
#         print("Disable success")
#     except subprocess.CalledProcessError as e:
#         print(f"Error occurred while disabling device: {e}")

# Locate and Auto-Click the Capture Button
def click_capture():
    global capture_button_located, capture_failed_located, capture_success_located, elapsed_time, timeout, interval

    save_button.configure(state="disabled")
    # clear_button.configure(state="disabled")

    while capture_button_located is None and elapsed_time < timeout:
        time.sleep(1)
        try:
            capture_button_located = pyautogui.locateOnScreen(resource_path("resources/capture_button.png"), confidence = 0.8)

            if capture_button_located:
                pyautogui.click(capture_button_located)
                print("Capture Button clicked!")

                time.sleep(1)
                locate_result = 0
                timeout_result = 10

                # while (capture_failed_located is None or capture_success_located is None) and locate_result < timeout_result:
                try:
                    capture_failed_located = pyautogui.locateOnScreen(resource_path("resources/capture_failed.png"), confidence=0.8)

                    if capture_failed_located:
                        print("Capture Failed!")
                        retry = pyautogui.confirm("Capture Failed. Try again?", buttons=["Yes","No"])
                        if retry == "Yes":
                            print("Trying again...")
                            capture_button_located = None
                            elapsed_time = 0
                            capture_failed_located = None
                            locate_result = 0
                        else:
                            print("Giving Up...")
                            save_button.configure(state="normal")
                            # clear_button.configure(state="normal")
                            capture_button_located = None
                            elapsed_time = 0
                            capture_failed_located = None
                            locate_result = 0
                            break
                    
                    locate_result += 1
                except pyautogui.ImageNotFoundException:
                    print("Image Not Found!: FAILED")
                    
                try:
                    capture_success_located = pyautogui.locateOnScreen(resource_path("resources/capture_success.png"), confidence=0.8)

                    if capture_success_located:
                        print("Capture Success!")
                        capture_button_located = None
                        elapsed_time = 0
                        capture_success_located = None
                        locate_result = 0
                        # popup = ctk.CTkToplevel(sl_window)
                        sl_window.after(0, shorted_window)
                        break
    
                except pyautogui.ImageNotFoundException:
                    print("Image Not Found!: SUCCESS")

                elapsed_time = 0
                # break

        except pyautogui.ImageNotFoundException:
            print("Image Not Found!")
    
        elapsed_time += interval

    if elapsed_time >= timeout:
        save_button.configure(state="normal")
        capture_button_located = None
        elapsed_time = 0
        capture_success_located = None
        locate_result = 0
        pyautogui.alert("Please try again. \nMake sure that only the ShortCamII application is open...", "Error!")
        print("Capture Button not found!")

# Uploading the Captured Image to the Database
def find_folder_directory (start_dir="C:/"):
    target_suffix = os.path.join("ShortCam II", "Record")

    for root, dirs, files in os.walk(start_dir):
        if root.endswith(target_suffix):
            return root
    return None

def get_latest_image(folder):
    image_files = [f for f in os.listdir(folder) if f.lower().endswith(('.jpg', '.jpeg', '.png'))]
    if not image_files:
        return None
    latest_file = max(image_files, key=lambda f: os.path.getctime(os.path.join(folder, f)))
    return os.path.join(folder, latest_file)
          
#####################################################
########### SQL Database Connection Part ############
#####################################################

# Establish SQL Connection
try:
    sql_connection = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};'
                                    'SERVER=' + server + ';'
                                    'DATABASE=' + database + ';'
                                    'UID=' + username + ';'
                                    'PWD=' + password + ';')
    print("Connected to SQL Server successfully")
except Exception as e:
    print("Error while connecting to SQL Server", e)
    pyautogui.alert("Error while connecting to SQL Server. \n\n\nPlease check if you are connected to Wistron network.")

# Create a cursor to interact with the Database
sql_cursor = sql_connection.cursor()
    
# def fetch_data():
#     global query_result
#     # Perform SQL Query to check if the SN exists
#     sql_cursor.execute("SELECT SN, Model, Failure_Description FROM zebrawip WHERE SN = ?", input_sn)

#     # Fetch the result (fetchone: single row, fetchall: multiple row)
#     row = sql_cursor.fetchone()

#     if row:
#         query_result = (row[0], row[1], row[2])
#         print(f"Data Found : {query_result}")
#     else:
#         query_result=None
#         print("No data found for that SN.")

# def close_sql_connection():
#     sql_cursor.close()
#     sql_connection.close()
#     print("SQL Server Disconnected!")

#####################################################
###### Customtkinter UI Design & Function Part ######
#####################################################

### CREATE CUSTOMTKINTER WINDOW ###
sl_window = ctk.CTk()
sl_window.geometry("370x85")
sl_window.overrideredirect(True)  # Removes title bar
sl_window.configure(fg_color="PaleTurquoise1")
sl_window.lift()
sl_window.attributes("-topmost", True)

# Pop-up Window for Details
def shorted_window():
    save_button.configure(state="disabled")
    # clear_button.configure(state="disabled")
    popup = ctk.CTkToplevel()
    width = 800
    height = 450
    screen_width = popup.winfo_screenwidth()
    screen_height = popup.winfo_screenheight()
    x = int((screen_width/2) - (width/2))
    y = int((screen_height/2) - (height/2))
    popup.geometry(f"{width}x{height}+{x}+{y}")
    popup.title("Shorted Component Detected")
    popup.configure(fg_color="PaleTurquoise1")
    popup.resizable(False, False)
    popup.attributes("-topmost", 1)
    
    def disable_popup_widgets():
        sn_entry.configure(state="disabled")
        model_entry.configure(state="disabled")
        issue_entry.configure(state="disabled")
        PN_entry.configure(state="disabled")
        Part_Type_entry.configure(state="disabled")
        helpful_checkbox.configure(state="disabled")
        nothelpful_checkbox.configure(state="disabled")
        save_component_button.configure(state="disabled")
        cancel_component_button.configure(state="disabled")

    def enable_popup_widgets():
        sn_entry.configure(state="normal")
        model_entry.configure(state="normal")
        issue_entry.configure(state="normal")
        PN_entry.configure(state="normal")
        Part_Type_entry.configure(state="normal")
        helpful_checkbox.configure(state="normal")
        nothelpful_checkbox.configure(state="normal")
        save_component_button.configure(state="normal")
        cancel_component_button.configure(state="normal")

    temp_label = None
    def show_uploading_status():
        global temp_label
        temp_label = ctk.CTkLabel(popup, text="Uploading to database...", text_color="red",
                                  font=("Arial", 20))
        temp_label.place(x=480, y=390)
        disable_popup_widgets()

    def get_form_values():
        fields = [
            sn_entry.get().strip(),
            model_entry.get().strip(),
            issue_entry.get().strip(),
            PN_entry.get().strip(),
            Part_Type_entry.get().strip()
        ]
        
        if helpful_checkbox.get():
            feedback_result = 1
        elif nothelpful_checkbox.get():
            feedback_result = 0
        else:
            feedback_result = None

        return fields, feedback_result

    def insert_data():
        global temp_label
        fields, feedback_result = get_form_values()
        image_data = None
        image_filename = None

         # Try to find the latest image
        folder = find_folder_directory("C:/")
        if folder:
            latest_image = get_latest_image(folder)
            if latest_image:
                with open(latest_image, 'rb') as img_file:
                    image_data = img_file.read()
                image_filename = os.path.basename(latest_image)
                print("Found and loaded image:", latest_image)
            else:
                print("No image found in folder.")
        else:
            print("'ShortCam II\\Record' folder not found.")


        insert_query = "INSERT INTO analysis (SN, Model, MB_Issue, PN_Component, Part_Type, Result, File_Name, Reference_Picture) VALUES (?, ?, ?, ?, ?, ?, ?, ?)"
        values = fields + [feedback_result, image_filename, image_data]


        try:
            sql_cursor.execute(insert_query, values)
            sql_connection.commit()
            print("Data Uploaded to Analysis Table")
            return True
        except Exception as e:
            pyautogui.alert("Error in uploading to database! Please try again. \n\n\nIf error persist, please contact CBE100.")
            print("Upload failed:", e)

            # Only destroy label if it exists
            if temp_label:
                try:
                    temp_label.destroy()
                except:
                    pass
                
            enable_popup_widgets()
            return False

    def background_task():
        success = insert_data()
        if success:
            popup.after(0, hide_uploading_status)

    def hide_uploading_status():
        global temp_label
        if temp_label:
            temp_label.destroy()
        close_popup()
        pyautogui.alert("Saved Successfully!", "Overheating Component Details")

    def save_data():
        show_uploading_status()  # Show label immediately on main thread
        threading.Thread(target=background_task).start()
        # clear_sn_box()
        # disable_thermalcam(Camera_ID)
        # close_sql_connection()
        
    def cancel_data():
        popup.destroy()
        save_button.configure(state="normal")
        # clear_button.configure(state="normal")

    def close_popup():
        save_button.configure(state="normal")
        # clear_button.configure(state="normal")
        popup.destroy()

    def update_checkboxes():
        if helpful_checkbox.get():
            nothelpful_checkbox.configure(state="disabled")
        elif nothelpful_checkbox.get():
            helpful_checkbox.configure(state="disabled")
        else:
            helpful_checkbox.configure(state="normal")
            nothelpful_checkbox.configure(state="normal")

        check_entries()

    # # Retrieve SN from entry box
    # def get_sn(event=None):
    #     global input_sn
    #     input_sn = sn_entry.get()

    #     if len(input_sn) == 24 and sn_entry.cget("state") != 'readonly':
    #         # enable_thermalcam(Camera_ID)
    #         print(f"Input SN: {input_sn}")
            
    #         # sn_entry.configure(state="readonly")  # Keeps text visible & selectable
    #         fetch_data()

    def check_entries(event=None):
        fields = [
        sn_entry.get().strip(),
        model_entry.get().strip(),
        issue_entry.get().strip(),
        PN_entry.get().strip(),
        Part_Type_entry.get().strip()
        ]
        
        checkbox_selected = helpful_checkbox.get() or nothelpful_checkbox.get()

        if all(fields) and checkbox_selected:
            save_component_button.configure(state="normal")
        else:
            save_component_button.configure(state="disabled")

    # Pop-up label
    popup_label = ctk.CTkLabel(popup, text="Details of Overheating Component", 
                               justify="center", font=("Arial", 20, "bold"),
                               text_color="black")
    popup_label.place(x=225, y=20)
    
    # SN text
    sn_text = ctk.CTkLabel(popup, text="SN:", font=("Arial", 20), text_color="black")
    sn_text.place(x=85, y=73)

    # SN Entry Box
    sn_entry = ctk.CTkEntry(popup, placeholder_text="Scan SN then press 'Enter'", placeholder_text_color="gray",
                            width=270, height=37, font=("Arial", 18), 
                            text_color="black", fg_color="white",border_width=0)
    # sn_entry.bind("<KeyRelease>", get_sn)
    sn_entry.place(x=125, y=70)

    # Model text
    model_text = ctk.CTkLabel(popup, text="Model:", 
                              font=("Arial", 18), text_color="black")
    model_text.place(x=505, y=73)

    # Model Entry Box
    model_entry = ctk.CTkEntry(popup, width=120, height=37, font=("Arial", 20), 
                               text_color="black", fg_color="white", border_width=0)
    model_entry.place(x=565, y=70)

    # Failure Description Entry Configuration
    issue_label = ctk.CTkLabel(popup, text="Mainboard Issue: ",
                                    font=("Arial", 18), text_color="black")
    issue_label.place(x=35, y=150)

    # Failure Description Entry Configuration
    issue_entry = ctk.CTkEntry(popup, width=390, height=30, 
                                   font=("Arial", 20), text_color="black",
                                   fg_color="white", border_width=0)
    issue_entry.place(x=350, y=145)

    # PN Entry Configuration
    PN_label = ctk.CTkLabel(popup, text="PN of Detected Component: ", width=20,
                            font=("Arial", 18), text_color="black")
    PN_label.place(x=35, y=207)

    PN_entry = ctk.CTkEntry(popup, width=390, height=30, 
                            font=("Arial", 20), text_color="black",
                            fg_color="white", border_width=0)
    PN_entry.place(x=350,y=205)
    PN_entry.bind("<KeyRelease>", check_entries)

    # Part Type Entry Configuration
    Part_Type_label = ctk.CTkLabel(popup, text="Part Type of Detected Component: ", width=20,
                                    font=("Arial", 18), text_color="black")
    Part_Type_label.place(x=35, y=264)

    Part_Type_entry = ctk.CTkEntry(popup, width=390, height=30, 
                                   font=("Arial", 20), text_color="black",
                                   fg_color="white", border_width=0)
    Part_Type_entry.place(x=350, y=265)
    Part_Type_entry.bind("<KeyRelease>", check_entries)

    # Helpful Checkbox Configuration
    helpful_checkbox = ctk.CTkCheckBox(popup, text="Helpful", command=update_checkboxes,
                                       width=140, height=28, checkbox_width=25, checkbox_height=25,
                                       font=("Arial", 18), text_color="black", checkmark_color="green2",
                                       fg_color="white", hover_color="green2", border_color="black",
                                       corner_radius=0,)
    helpful_checkbox.place(x=50, y=320)
    # helpful_checkbox.bind("<KeyRelease>", check_entries)

    # Not Helpful Checkbox Configuration
    nothelpful_checkbox = ctk.CTkCheckBox(popup, text="Not Helpful", command=update_checkboxes,
                                       width=140, height=28, checkbox_width=25, checkbox_height=25,
                                       font=("Arial", 18), text_color="black", checkmark_color="green2",
                                       fg_color="white", hover_color="green2", border_color="black",
                                       corner_radius=0,)
    nothelpful_checkbox.place(x=50, y=365)
    # nothelpful_checkbox.bind("<KeyRelease>", check_entries)

    # Save Details Button to Add to Record
    save_component_button = ctk.CTkButton(popup, text="Save", command=save_data,
                                          width=90, height=40, corner_radius=30,
                                          font=("Arial",18), text_color="black",
                                          fg_color="white",hover_color="green2", 
                                          border_color="black", border_width=2)
    save_component_button.place(x=480, y=335)
    save_component_button.configure(state="disabled")

    # Cancel Details Button
    cancel_component_button = ctk.CTkButton(popup, text="Cancel", command=cancel_data,
                                          width=90, height=40, corner_radius=30,
                                          font=("Arial",18), text_color="black",
                                          fg_color="white",hover_color="red2", 
                                          border_color="black", border_width=2)
    cancel_component_button.place(x=620, y=335)

    # popup.mainloop()

    popup.protocol("WM_DELETE_WINDOW", close_popup)

########## BUTTON CONFIGURATIONS ##########
        
# # Clear SN box
# def clear_sn_box():
#     save_button.configure(state="disabled")
#     # clear_button.place_forget()
#     sn_entry.configure(state="normal")
#     sn_entry.delete(0, "end")
#     # disable_thermalcam(Camera_ID)

# Load Logo Image
logo_image = ctk.CTkImage(light_image=Image.open(resource_path("resources/SL logo.png")), size=(78, 85))
logo_image_label = ctk.CTkLabel(sl_window, image=logo_image, text="")
logo_image_label.place(x=3, y=0)

# Short Locator Label
sl_label = ctk.CTkLabel(sl_window, text="Short Locator Logs",
                        font=("Arial", 18, "bold"), text_color="black")
sl_label.place(x=135, y=7)

# Save Button
save_button = ctk.CTkButton(sl_window, text="Save", command=click_capture,
                            corner_radius=10, border_color="black", font=("Arial", 15, "bold"),
                            width=90, height=35, border_width=2,
                            fg_color="white", text_color="black", hover_color="green2")
save_button.place(x=180, y=43)
save_button.configure(state="normal")

# # Logs Button
# logs_button = ctk.CTkButton(sl_window, text="View Logs", 
#                             corner_radius=10, border_color="black", font=("Arial", 15, "bold"),
#                             width=70, height=35, border_width=2,
#                             fg_color="white", text_color="black", hover_color="green2")
# logs_button.place(x=110, y=43)
# logs_button.configure(state="normal")

# # Clear Button
# clear_button = ctk.CTkButton(sl_window, command=clear_sn_box,
#                       text="x", text_color="black", font=("Arial", 15, "bold"),
#                       corner_radius=100, border_color="black", width=10, height=10,
#                       fg_color="transparent", hover_color="red")

### WINDOW POSITION TRACKING ###
def follow_app():
    if ShortCam_win:
        try:
            rect = win32gui.GetWindowRect(ShortCam_win)
            x, y = rect[0], rect[1]  # Top-left Corner
            sl_window.geometry(f"+{x}+{y + 25}")  # Offset for visibility
        except:
            print("Exiting Tkinter Window...")
            sl_window.destroy()
            return

    sl_window.after(10, follow_app)

### MONITOR IF SHORTCAM II IS CLOSED ###
def monitor_app():
    if not is_process_running("ShortCam II.exe"):
        print("ShortCam II closed. Closing UI...")
        sl_window.destroy()
    else:
        sl_window.after(1000, monitor_app)  # Check every second

# Start tracking ShortCam II movement and monitor closure
follow_app()
monitor_app()

sl_window.mainloop()