import customtkinter as ctk
from PIL import Image

root = ctk.CTk()
root.title("Short Locator Log Input")

# Desired window size
window_width = 1050
window_height = 750

# Center the window on screen
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
center_x = int((screen_width / 2) - (window_width / 2))
center_y = int((screen_height / 2) - (window_height / 2))
root.geometry(f"{window_width}x{window_height}+{center_x}+{center_y}")
root.resizable(False, False)

# --- Configure the main grid layout ---
# This creates padding around the entire window content
root.grid_columnconfigure(0, weight=1)
root.grid_rowconfigure(5, weight=1) # Allow the middle section to expand if needed
root.grid_rowconfigure(9, weight=1) # Push buttons to the bottom

# --- Header Frame (for LOGO and SEE LOGS button) ---
header_frame = ctk.CTkFrame(root, fg_color="transparent")
header_frame.grid(row=0, column=0, padx=50, pady=(20, 10), sticky="ew")
header_frame.grid_columnconfigure(0, weight=1) # Makes the logo side expandable
header_frame.grid_columnconfigure(1, weight=0) # See logs button takes only its own space

# --- Placeholder for LOGO ---
# You can add your CTkImage label here
logo_label = ctk.CTkLabel(header_frame, text="LOG", font=ctk.CTkFont(size=36, weight="bold"))
logo_label.grid(row=0, column=0, sticky="w")

# --- SEE LOGS Button ---
see_logs_button = ctk.CTkButton(header_frame, text="SEE LOGS", width=120, height=40, font=ctk.CTkFont(size=14))
see_logs_button.grid(row=0, column=1, sticky="e")


# --- Content Frame (for the main form fields) ---
content_frame = ctk.CTkFrame(root, fg_color="transparent")
content_frame.grid(row=1, column=0, rowspan=8, padx=50, pady=20, sticky="nsew")
content_frame.grid_columnconfigure((0, 1), weight=1) # Allow both columns to take up space


# --- S/N and MODEL ---
serial_label = ctk.CTkLabel(content_frame, text="S/N:", font=ctk.CTkFont(size=16))
serial_label.grid(row=0, column=0, sticky="w")

# Using CTkEntry for single-line input
serial_entry = ctk.CTkEntry(content_frame, height=40, font=ctk.CTkFont(size=14))
serial_entry.grid(row=1, column=0, sticky="ew", pady=(5, 20), padx=(0, 20)) # Added padx for spacing

model_label = ctk.CTkLabel(content_frame, text="MODEL:", font=ctk.CTkFont(size=16))
model_label.grid(row=0, column=1, sticky="w", padx=(20, 0)) # Added padx for spacing

# Using CTkEntry for single-line input
model_entry = ctk.CTkEntry(content_frame, height=40, font=ctk.CTkFont(size=14))
model_entry.grid(row=1, column=1, sticky="ew", pady=(5, 20), padx=(20, 0)) # Added padx for spacing


# --- FAILURE DESCRIPTION ---
description_label = ctk.CTkLabel(content_frame, text="FAILURE DESCRIPTION:", font=ctk.CTkFont(size=16))
description_label.grid(row=2, column=0, columnspan=2, sticky="w", pady=(20, 5))

# Textbox is correct for multi-line descriptions
description_text_field = ctk.CTkTextbox(content_frame, height=120)
description_text_field.grid(row=3, column=0, columnspan=2, sticky="ew")


# --- OVERHEATING COMPONENT DETAILS ---
overheating_label = ctk.CTkLabel(content_frame, text="OVERHEATING COMPONENT DETAILS", font=ctk.CTkFont(size=18, weight="bold"))
overheating_label.grid(row=4, column=0, columnspan=2, sticky="w", pady=(40, 20))

pn_label = ctk.CTkLabel(content_frame, text="PN of Detected Component:", font=ctk.CTkFont(size=14))
pn_label.grid(row=5, column=0, columnspan=2, sticky="w", pady=(0, 5))

# Using CTkEntry for single-line input
pn_entry = ctk.CTkEntry(content_frame, height=40, font=ctk.CTkFont(size=14))
pn_entry.grid(row=6, column=0, columnspan=2, sticky="ew")

part_label = ctk.CTkLabel(content_frame, text="Part Type of Detected Component:", font=ctk.CTkFont(size=14))
part_label.grid(row=7, column=0, columnspan=2, sticky="w", pady=(20, 5))

# Using CTkEntry for single-line input
part_entry = ctk.CTkEntry(content_frame, height=40, font=ctk.CTkFont(size=14))
part_entry.grid(row=8, column=0, columnspan=2, sticky="ew")


# --- Buttons Frame (to group SAVE and CANCEL) ---
button_frame = ctk.CTkFrame(root, fg_color="transparent")
button_frame.grid(row=10, column=0, pady=(20, 40)) # Pushes buttons down with pady

save_button = ctk.CTkButton(button_frame, text="SAVE", width=140, height=40)
save_button.grid(row=0, column=0, padx=10)

cancel_button = ctk.CTkButton(button_frame, text="CANCEL", width=140, height=40)
cancel_button.grid(row=0, column=1, padx=10)


root.mainloop()