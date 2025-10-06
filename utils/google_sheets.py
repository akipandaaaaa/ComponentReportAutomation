"""
Google Sheets API wrapper with enhanced data validation detection
Handles hidden sheets, cross-sheet references, and dependent dropdowns
"""
import gspread
from google.oauth2.service_account import Credentials
from utils.config import SCOPES, CREDENTIALS_FILE
import os
import time
import re

class GoogleSheetsManager:
    def __init__(self):
        self.client = None
        self.current_sheet = None
        self.current_worksheet = None
        self.connected = False
        self.spreadsheet_id = None
        self.credentials = None
        
    def connect(self):
        """Connect to Google Sheets using service account"""
        try:
            if not os.path.exists(CREDENTIALS_FILE):
                raise FileNotFoundError(
                    "credentials.json not found. Please add your service account credentials."
                )
            
            self.credentials = Credentials.from_service_account_file(
                CREDENTIALS_FILE, 
                scopes=SCOPES
            )
            self.client = gspread.authorize(self.credentials)
            self.connected = True
            return True, "Connected successfully"
        except FileNotFoundError as e:
            return False, str(e)
        except Exception as e:
            return False, f"Connection failed: {str(e)}"
    
    def open_spreadsheet(self, url):
        """Open a spreadsheet by URL"""
        try:
            if not self.connected:
                return False, "Not connected. Please connect first."
            
            self.current_sheet = self.client.open_by_url(url)
            self.spreadsheet_id = url.split('/d/')[1].split('/')[0]
            return True, f"Opened: {self.current_sheet.title}"
        except Exception as e:
            return False, f"Failed to open spreadsheet: {str(e)}"
    
    def get_worksheet_names(self, include_hidden=False):
        """
        Get all worksheet/tab names
        Args:
            include_hidden: If True, includes hidden sheets
        """
        try:
            if not self.current_sheet:
                return []
            
            worksheets = self.current_sheet.worksheets()
            
            if include_hidden:
                return [ws.title for ws in worksheets]
            else:
                # Filter out hidden sheets
                visible = []
                for ws in worksheets:
                    if not ws.isSheetHidden:
                        visible.append(ws.title)
                return visible if visible else [ws.title for ws in worksheets]
                
        except Exception as e:
            print(f"Error getting worksheets: {e}")
            return []
    
    def get_worksheet(self, name):
        """Get specific worksheet by name (works with hidden sheets too)"""
        try:
            self.current_worksheet = self.current_sheet.worksheet(name)
            return self.current_worksheet
        except Exception as e:
            print(f"Error getting worksheet '{name}': {e}")
            return None
    
    def parse_range_reference(self, range_str):
        """
        Parse a range reference and extract sheet name and range
        Handles formulas like =Backend!$AC$2:$AC
        Returns: (sheet_name, range_part)
        """
        if not range_str:
            return None, None
        
        # Remove leading equals sign if present
        if range_str.startswith('='):
            range_str = range_str[1:]
        
        # Clean up the string
        range_str = range_str.strip()
        dollar_sign = '$'
        range_str = range_str.replace(dollar_sign, '')
        
        # Check for sheet reference (contains !)
        if '!' in range_str:
            parts = range_str.split('!', 1)
            sheet_name = parts[0].strip("'\"")
            range_part = parts[1]
            
            # Fix incomplete ranges like AC2:AC (missing end row)
            if ':' in range_part:
                start, end = range_part.split(':', 1)
                # If end part has no digits, it is incomplete
                if end and not any(c.isdigit() for c in end):
                    # Extract column letter from end
                    col_letter = ''.join(c for c in end if c.isalpha())
                    old_range = range_part
                    range_part = start + ':' + col_letter + '1000'
                    print(f"[DEBUG] Fixed incomplete range: {old_range} -> {range_part}")
            
            return sheet_name, range_part
        else:
            return None, range_str
    
    def get_range_from_any_sheet(self, range_str, default_worksheet=None):
        """
        Get values from a range that may reference any sheet (including hidden)
        Handles formulas like =Backend!$AC$2:$AC
        """
        try:
            print(f"[DEBUG] Processing range: {range_str}")
            sheet_name, range_part = self.parse_range_reference(range_str)
            
            # Determine which worksheet to use
            if sheet_name:
                # Explicit sheet reference - get that sheet (even if hidden)
                target_worksheet = self.get_worksheet(sheet_name)
                if not target_worksheet:
                    print(f"[ERROR] Could not access sheet '{sheet_name}'")
                    return []
                print(f"[DEBUG] Reading from sheet: '{sheet_name}' (may be hidden)")
            else:
                # No sheet reference - use default
                target_worksheet = default_worksheet
                if not target_worksheet:
                    print("[ERROR] No worksheet specified and no default provided")
                    return []
            
            # Read the range
            print(f"[DEBUG] Reading range: {range_part}")
            values = self._read_range_safe(target_worksheet, range_part)
            
            # Flatten and clean
            result = []
            seen = set()  # Track duplicates
            for row in values:
                if isinstance(row, list):
                    for val in row:
                        if val and str(val).strip():
                            clean_val = str(val).strip()
                            if clean_val not in seen:
                                seen.add(clean_val)
                                result.append(clean_val)
                elif row and str(row).strip():
                    clean_val = str(row).strip()
                    if clean_val not in seen:
                        seen.add(clean_val)
                        result.append(clean_val)
            
            print(f"[DEBUG] Extracted {len(result)} unique values (removed {len(seen) - len(result)} duplicates)")
            return result
            
        except Exception as e:
            print(f"[ERROR] Error reading range '{range_str}': {e}")
            import traceback
            traceback.print_exc()
            return []
    
    def detect_data_validations(self, worksheet_name):
        """Detect all data validation rules in a worksheet"""
        try:
            from google.auth.transport.requests import AuthorizedSession
            
            worksheet = self.current_sheet.worksheet(worksheet_name)
            sheet_id = worksheet.id
            
            authed_session = AuthorizedSession(self.credentials)
            
            url = f"https://sheets.googleapis.com/v4/spreadsheets/{self.spreadsheet_id}"
            params = {'fields': 'sheets(properties,data.rowData.values.dataValidation)'}
            
            response = authed_session.get(url, params=params)
            
            if response.status_code != 200:
                print(f"API Error: {response.status_code}")
                return []
            
            data = response.json()
            validations = []
            
            for sheet in data.get('sheets', []):
                if sheet['properties']['sheetId'] == sheet_id:
                    sheet_data = sheet.get('data', [])
                    
                    for grid_data in sheet_data:
                        row_data = grid_data.get('rowData', [])
                        
                        for row_idx, row in enumerate(row_data):
                            values = row.get('values', [])
                            
                            for col_idx, cell in enumerate(values):
                                if 'dataValidation' in cell:
                                    validation = cell['dataValidation']
                                    condition = validation.get('condition', {})
                                    
                                    cell_address = f"{self._col_num_to_letter(col_idx + 1)}{row_idx + 1}"
                                    validation_type = condition.get('type', 'UNKNOWN')
                                    range_ref = None
                                    
                                    if validation_type == 'ONE_OF_RANGE':
                                        condition_values = condition.get('values', [])
                                        for val in condition_values:
                                            user_value = val.get('userEnteredValue', '')
                                            if user_value:
                                                range_ref = user_value
                                                print(f"[DEBUG] Found range reference: {range_ref}")
                                    
                                    validations.append({
                                        'cell': cell_address,
                                        'type': validation_type,
                                        'range': range_ref,
                                        'worksheet': worksheet,
                                        'referenced_sheet': self.parse_range_reference(range_ref)[0] if range_ref else None
                                    })
            
            return validations
            
        except Exception as e:
            print(f"Error detecting data validations: {e}")
            import traceback
            traceback.print_exc()
            return []
    
    def read_dropdown_values_from_cell(self, worksheet, cell_address, sheet_name):
        """
        Read dropdown values from a cell by detecting its data validation
        Now supports hidden sheets, cross-sheet references, and formulas
        """
        try:
            print(f"[DEBUG] ========================================")
            print(f"[DEBUG] Reading dropdown from cell {cell_address} in sheet '{sheet_name}'")
            
            # Detect validations for this sheet
            validations = self.detect_data_validations(sheet_name)
            print(f"[DEBUG] Found {len(validations)} total validations in sheet")
            
            # Show all validations for debugging
            for i, val in enumerate(validations, 1):
                print(f"[DEBUG] Validation {i}: cell={val['cell']}, type={val['type']}, range={val.get('range')}")
            
            # Find the validation for our target cell
            target_validation = None
            for val in validations:
                if val['cell'] == cell_address:
                    target_validation = val
                    print(f"[DEBUG] Found target validation:")
                    print(f"[DEBUG]   - Cell: {val['cell']}")
                    print(f"[DEBUG]   - Type: {val['type']}")
                    print(f"[DEBUG]   - Range: {val.get('range')}")
                    print(f"[DEBUG]   - Referenced sheet: {val.get('referenced_sheet')}")
                    break
            
            if not target_validation:
                print(f"[ERROR] No data validation found for cell {cell_address}")
                print(f"[DEBUG] Available cells with validation: {[v['cell'] for v in validations]}")
                return []
            
            range_ref = target_validation.get('range')
            if not range_ref:
                print(f"[ERROR] No range reference found in validation for {cell_address}")
                return []
            
            print(f"[DEBUG] Attempting to read values from: {range_ref}")
            
            # Use the new method that handles hidden sheets and formulas
            values = self.get_range_from_any_sheet(range_ref, worksheet)
            
            if values:
                print(f"[DEBUG] Successfully read {len(values)} values")
                print(f"[DEBUG] First 5 values: {values[:5]}")
            else:
                print(f"[DEBUG] No values returned")
            
            print(f"[DEBUG] ========================================")
            return values
            
        except Exception as e:
            print(f"[ERROR] Error reading dropdown from cell {cell_address}: {e}")
            import traceback
            traceback.print_exc()
            return []
    
    def get_dropdown_values(self, worksheet, range_str):
        """
        Legacy method - now redirects to get_range_from_any_sheet
        Kept for backwards compatibility
        """
        return self.get_range_from_any_sheet(range_str, worksheet)
    
    def _read_range_safe(self, worksheet, range_str):
        """
        Safely read a range with multiple fallback strategies
        Handles incomplete ranges and grid limit errors
        """
        try:
            # Try direct read first
            return worksheet.get(range_str)
        except Exception as e:
            error_msg = str(e).lower()
            
            # Handle exceeds grid limits error
            if "exceeds grid limits" in error_msg or "out of bounds" in error_msg:
                print(f"Range exceeds limits, trying fallback strategies...")
                
                # Strategy 1: If it is a column range, limit to reasonable rows
                if ':' in range_str:
                    start, end = range_str.split(':', 1)
                    
                    # Full column reference (e.g., A:A)
                    if not any(c.isdigit() for c in start) and not any(c.isdigit() for c in end):
                        col_letter = ''.join(c for c in end if c.isalpha())
                        limited_range = col_letter + '1:' + col_letter + '1000'
                        print(f"Trying limited range: {limited_range}")
                        try:
                            return worksheet.get(limited_range)
                        except:
                            pass
                    
                    # Incomplete end (e.g., A1:A without row number)
                    elif not any(c.isdigit() for c in end):
                        col_letter = ''.join(c for c in end if c.isalpha())
                        limited_range = start + ':' + col_letter + '1000'
                        print(f"Trying limited range: {limited_range}")
                        try:
                            return worksheet.get(limited_range)
                        except:
                            pass
                
                # Strategy 2: Get all values and filter by column
                print("Trying to get all values and filter...")
                try:
                    all_values = worksheet.get_all_values()
                    # Extract column from range
                    if ':' in range_str:
                        start_cell = range_str.split(':')[0]
                        col_letter = ''.join(c for c in start_cell if c.isalpha())
                        col_index = self._col_letter_to_num(col_letter) - 1
                        
                        filtered = []
                        for row in all_values:
                            if col_index < len(row) and row[col_index]:
                                filtered.append([row[col_index]])
                        return filtered
                except Exception as e2:
                    print(f"Fallback strategy failed: {e2}")
            
            # If all strategies fail, return empty
            print(f"Could not read range: {range_str}")
            return []
    
    def _col_num_to_letter(self, n):
        """Convert column number to letter (1 -> A, 2 -> B, etc.)"""
        string = ""
        while n > 0:
            n, remainder = divmod(n - 1, 26)
            string = chr(65 + remainder) + string
        return string
    
    def _col_letter_to_num(self, letters):
        """Convert column letter to number (A -> 1, B -> 2, etc.)"""
        num = 0
        for char in letters.upper():
            num = num * 26 + (ord(char) - ord('A') + 1)
        return num
    
    def get_cell_value(self, worksheet, cell):
        """Get value from a specific cell"""
        try:
            return worksheet.acell(cell).value
        except Exception as e:
            print(f"Error getting cell value: {e}")
            return None
    
    def set_cell_value(self, worksheet, cell, value):
        """Set value in a specific cell"""
        try:
            worksheet.update_acell(cell, value)
            time.sleep(2)
            return True
        except Exception as e:
            print(f"Error setting cell value: {e}")
            return False
    
    def find_last_row_with_data(self, worksheet, column_letter, start_row):
        """Find the last row with data in a specific column"""
        try:
            col_num = self._col_letter_to_num(column_letter)
            column_values = worksheet.col_values(col_num)
            
            for i in range(len(column_values) - 1, start_row - 2, -1):
                if i >= 0 and column_values[i] and str(column_values[i]).strip():
                    return i + 1
            
            return start_row
        except Exception as e:
            print(f"Error finding last row: {e}")
            return start_row
    
    def export_range_as_pdf(self, worksheet_name, cell_range, output_path):
        """Export a specific range as PDF"""
        try:
            if not self.spreadsheet_id:
                return False, "No spreadsheet opened"
            
            worksheet = self.current_sheet.worksheet(worksheet_name)
            sheet_id = worksheet.id
            
            base_url = f"https://docs.google.com/spreadsheets/d/{self.spreadsheet_id}"
            export_url = f"{base_url}/export?format=pdf&gid={sheet_id}&range={cell_range}"
            
            from google.auth.transport.requests import AuthorizedSession
            authed_session = AuthorizedSession(self.credentials)
            
            response = authed_session.get(export_url)
            
            if response.status_code == 200:
                with open(output_path, 'wb') as f:
                    f.write(response.content)
                return True, f"Exported to {output_path}"
            else:
                return False, f"Export failed: {response.status_code}"
                
        except Exception as e:
            return False, f"Export error: {str(e)}"
    
    def export_range_as_excel(self, worksheet_name, cell_range, output_path):
        """Export a specific range as Excel"""
        try:
            worksheet = self.current_sheet.worksheet(worksheet_name)
            data = worksheet.get(cell_range)
            
            import pandas as pd
            df = pd.DataFrame(data)
            df.to_excel(output_path, index=False, header=False)
            
            return True, f"Exported to {output_path}"
        except Exception as e:
            return False, f"Export error: {str(e)}"
    
    def export_range_as_csv(self, worksheet_name, cell_range, output_path):
        """Export a specific range as CSV"""
        try:
            worksheet = self.current_sheet.worksheet(worksheet_name)
            data = worksheet.get(cell_range)
            
            import csv
            with open(output_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerows(data)
            
            return True, f"Exported to {output_path}"
        except Exception as e:
            return False, f"Export error: {str(e)}"

# Global instance
sheets_manager = GoogleSheetsManager()