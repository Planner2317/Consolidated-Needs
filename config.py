import os
from datetime import datetime

# Define the file paths
FOLDER_PATH = r"C:\Users\Ahmed\Desktop\PythonLap\Gasket"
MASTER_PATH = os.path.join(FOLDER_PATH, "Master.xlsx")
RESPONSES_PATH = os.path.join(FOLDER_PATH, "Responses.xlsx")
OUTPUT_PATH = os.path.join(FOLDER_PATH, f"Gasket_Analysis_{datetime.now().strftime('%Y-%m-%d')}.xlsx")

# Constants - updated per the latest specification
CURRENT_DATETIME = "2025-04-22 12:55:05"  # Updated from your input
CURRENT_USER = "Planner2317"

# Status categories with color coding
STATUS = {
    'ACCEPTABLE': {'color': 'CCFFCC'},      # Green for exact match
    'MODERATE_DEVIATION': {'color': 'FFFF99'},  # Yellow for difference 1-3
    'HIGH_DEVIATION': {'color': 'FFCCCC'},   # Red for difference > 3
    'LOW_REQUEST': {'color': 'CCE5FF'},      # Blue for difference < 0
    'COVERED_BY_STOCK': {'color': '9999FF'}, # Purple for items covered by existing stock
    'COVERED_BY_ORDERS': {'color': 'FF99CC'} # Pink for items covered by pending orders
}

# Master file columns
INVENTORY_COLUMNS = [
    'Item Code', 'Description', 'Annual Forecast',
    'Stock Qty', 'Open PRs Total 24 Months', 'Open POs Total 24 Months', 'Pr Not Confirmed 24 Months'
]