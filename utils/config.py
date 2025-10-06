"""
Configuration settings for the automation application
"""
import os
import sys

# Determine if running as exe or script
if getattr(sys, 'frozen', False):
    # Running as compiled exe
    BASE_DIR = os.path.dirname(sys.executable)
else:
    # Running as script
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Application Settings
APP_NAME = "Component Report Automation Hub"
APP_VERSION = "1.0.0"
WINDOW_SIZE = "800x600"

# Paths
DOWNLOADS_DIR = os.path.join(BASE_DIR, "downloads")
CREDENTIALS_FILE = os.path.join(BASE_DIR, "credentials.json")

# Ensure downloads directory exists
os.makedirs(DOWNLOADS_DIR, exist_ok=True)

# Google Sheets Settings
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

# UI Colors (CustomTkinter themes)
COLORS = {
    "primary": "#1f6aa5",
    "success": "#2fa572",
    "warning": "#e6a23c",
    "danger": "#f56c6c",
    "info": "#909399"
}