import time
import pyautogui
from openpyxl import load_workbook
from openpyxl.drawing.image import Image

# Wait for a few seconds to switch to the Chrome tab with Snowflake
time.sleep(5)

# Capture screenshot of the Chrome tab
screenshot = pyautogui.screenshot()

# Save the screenshot to a file
img_path = 'screenshot.png'
screenshot.save(img_path)

# Load the existing Excel file
excel_file = r"C:\Users\patil\OneDrive\Documents\Demo_sheet.xlsx"
wb = load_workbook(excel_file)

# Select the active sheet
sheet = wb['Sheet3']

# Add the screenshot to the Excel sheet
img = Image(img_path)
# Assuming cell coordinates for pasting the image
# Example:
sheet.add_image(img, 'A1')

# Save the changes to the Excel file
wb.save(excel_file)
