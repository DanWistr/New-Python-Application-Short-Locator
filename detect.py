import pyautogui
APP_NAME = "ShortCam II"

try:
    app_window = pyautogui.getWindowsWithTitle(APP_NAME)[0]  # Get the first matching window
    app_window.activate()  # Bring the app window to the foreground
    print(f"Application window '{APP_NAME}' activated.")
except IndexError:
    print(f"Error: No window with title '{APP_NAME}' found.")
    exit()