import openpyxl
import os
from openpyxl import Workbook

# Define the XPaths, their descriptions, and the pages where they are used
xpaths_info = [
    {"xpath": "//input[@id='js-search-autocomplete']", "purpose": "Search input box", "pages": ["Home Page"]},
    {"xpath": "//ul[@id='js-search-items']", "purpose": "List containing search suggestions", "pages": ["Home Page"]},
    {"xpath": ".//li[contains(@class, 'google-auto-suggestion-list')]", "purpose": "Search suggestion items", "pages": ["Home Page"]},
    {"xpath": "//input[@id='js-date-range-display']", "purpose": "Date range input field", "pages": ["Home Page"]},
    {"xpath": "//td[contains(@class, 'datepicker__month-day--valid') and text()='${check_in_day}']", "purpose": "Check-in date picker cell", "pages": ["Home Page"]},
    {"xpath": "//td[contains(@class, 'datepicker__month-day--valid') and text()='${check_out_day}']", "purpose": "Check-out date picker cell", "pages": ["Home Page"]},
    {"xpath": "//button[@id='js-date-select']", "purpose": "Date selection confirmation button", "pages": ["Home Page"]},
    {"xpath": "//div[@id='js-btn-search']", "purpose": "Search button", "pages": ["Home Page"]},
    {"xpath": "//div[contains(@class, 'property-tiles')]//div[contains(@class, 'title')]//a", "purpose": "Property tiles anchor links", "pages": ["Home Page", "Refine Page"]},
    {"xpath": "//div[@id='js-date-available'] | //div[@id='js-date-unavailable']", "purpose": "Availability indicators", "pages": ["Home Page", "Refine Page"]},
    {"xpath": "//div[@id='js-date-available']", "purpose": "Available status indicator", "pages": ["Check Availability Page"]},
    {"xpath": "//div[@id='js-date-unavailable']", "purpose": "Unavailable status indicator", "pages": ["Check Availability Page"]}
]

# Create a new Excel workbook and select the active sheet
workbook = Workbook()
sheet = workbook.active
sheet.title = "XPath Information"

# Populate the first row with XPaths
for col_num, xpath_info in enumerate(xpaths_info, start=1):
    sheet.cell(row=1, column=col_num, value=xpath_info["xpath"])

    # Combine purpose and page info in the second row
    details = f"Purpose: {xpath_info['purpose']} | Used in: {', '.join(xpath_info['pages'])}"
    sheet.cell(row=2, column=col_num, value=details)

# Define the path to the data folder
data_folder_path = os.path.join("data")
# Ensure the data folder exists
os.makedirs(data_folder_path, exist_ok=True)

# Save the workbook to a file in the data folder
excel_file_path = os.path.join(data_folder_path, "xpaths_info.xlsx")
workbook.save(excel_file_path)
print(f"Excel file '{excel_file_path}' created successfully.")
