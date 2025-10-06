"""
Component Report Download automation screen
Simplified flow: User sets B3 manually → System processes B6 values automatically
Handles dependent dropdowns with sentinel monitoring
"""
import customtkinter as ctk
from tkinter import filedialog
from utils.google_sheets import sheets_manager
from utils.config import DOWNLOADS_DIR
import os
from datetime import datetime
import threading
import time

class ComponentReportScreen(ctk.CTkFrame):
    def __init__(self, parent, on_back):
        super().__init__(parent)
        self.on_back = on_back
        self.is_running = False
        self.component_dropdown_cell = "B6"
        self.menu_display_cell = "B3"
        self.component_values = []
        self.current_menu_value = None
        self.failed_components = []  # Track failed components
        
        # Header
        header = ctk.CTkFrame(self, height=60)
        header.pack(fill="x", padx=20, pady=(10, 0))
        header.pack_propagate(False)
        
        back_btn = ctk.CTkButton(header, text="← Back", command=self.go_back, width=80, height=35)
        back_btn.pack(side="left", pady=10)
        
        title = ctk.CTkLabel(header, text="Component Report Download", font=("Arial", 20, "bold"))
        title.pack(side="left", padx=20, pady=10)
        
        help_btn = ctk.CTkButton(header, text="?", command=self.show_help, width=35, height=35, font=("Arial", 16, "bold"))
        help_btn.pack(side="right", pady=10)
        
        # Scrollable content
        self.scrollable = ctk.CTkScrollableFrame(self)
        self.scrollable.pack(fill="both", expand=True, padx=20, pady=10)
        
        # CONNECTION SECTION
        self.create_section_header("STEP 1: GOOGLE SHEETS CONNECTION")
        
        instruction_label = ctk.CTkLabel(
            self.scrollable,
            text="⚠ Before starting: Manually set B3 (Menu) value in your Google Sheet first",
            text_color="#e6a23c",
            anchor="w",
            font=("Arial", 11, "bold")
        )
        instruction_label.pack(anchor="w", pady=(0, 10))
        
        url_frame = ctk.CTkFrame(self.scrollable, fg_color="transparent")
        url_frame.pack(fill="x", pady=10)
        
        ctk.CTkLabel(url_frame, text="Spreadsheet URL:", anchor="w").pack(anchor="w", pady=(0, 5))
        
        url_input_frame = ctk.CTkFrame(url_frame, fg_color="transparent")
        url_input_frame.pack(fill="x")
        
        self.url_entry = ctk.CTkEntry(url_input_frame, placeholder_text="Paste Google Sheets URL here...", height=35)
        self.url_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        self.test_btn = ctk.CTkButton(url_input_frame, text="Connect & Load", command=self.connect_and_load, width=120, height=35)
        self.test_btn.pack(side="right")
        
        self.status_label = ctk.CTkLabel(url_frame, text="Status: Not connected", anchor="w", text_color="gray")
        self.status_label.pack(anchor="w", pady=(5, 0))
        
        # MENU DISPLAY (Read-only)
        self.create_section_header("STEP 2: MENU CONFIGURATION")
        
        menu_frame = ctk.CTkFrame(self.scrollable, fg_color="transparent")
        menu_frame.pack(fill="x", pady=10)
        
        ctk.CTkLabel(menu_frame, text="Current Menu Selection (B3):", anchor="w").pack(anchor="w", pady=(0, 5))
        
        self.menu_display = ctk.CTkEntry(menu_frame, height=35, state="disabled")
        self.menu_display.pack(fill="x", pady=(0, 10))
        
        refresh_frame = ctk.CTkFrame(menu_frame, fg_color="transparent")
        refresh_frame.pack(fill="x")
        
        ctk.CTkLabel(
            refresh_frame,
            text="If you changed B3 in Google Sheets, click refresh:",
            anchor="w",
            text_color="gray"
        ).pack(side="left", expand=True)
        
        self.refresh_btn = ctk.CTkButton(
            refresh_frame,
            text="Refresh B6 Values",
            command=self.refresh_component_values,
            width=140,
            height=35,
            state="disabled"
        )
        self.refresh_btn.pack(side="right")
        
        # COMPONENT PREVIEW
        self.create_section_header("STEP 3: COMPONENT LIST (B6)")
        
        component_frame = ctk.CTkFrame(self.scrollable, fg_color="transparent")
        component_frame.pack(fill="x", pady=10)
        
        self.component_preview_label = ctk.CTkLabel(component_frame, text="", anchor="w", text_color="gray")
        self.component_preview_label.pack(anchor="w", pady=(0, 10))
        
        self.component_preview_box = ctk.CTkTextbox(component_frame, height=120)
        self.component_preview_box.pack(fill="x")
        
        # DATA RANGE SETTINGS
        self.create_section_header("STEP 4: DATA RANGE DETECTION SETTINGS")
        
        range_frame = ctk.CTkFrame(self.scrollable, fg_color="transparent")
        range_frame.pack(fill="x", pady=10)
        
        range_inputs = ctk.CTkFrame(range_frame, fg_color="transparent")
        range_inputs.pack(fill="x")
        
        # Row 1: Start cell, End column, Check column
        row1 = ctk.CTkFrame(range_inputs, fg_color="transparent")
        row1.pack(fill="x", pady=(0, 10))
        
        start_frame = ctk.CTkFrame(row1, fg_color="transparent")
        start_frame.pack(side="left", fill="x", expand=True, padx=(0, 5))
        ctk.CTkLabel(start_frame, text="Start Cell:", anchor="w").pack(anchor="w")
        self.start_entry = ctk.CTkEntry(start_frame, placeholder_text="A9", height=35)
        self.start_entry.insert(0, "A9")
        self.start_entry.pack(fill="x", pady=(5, 0))
        
        end_frame = ctk.CTkFrame(row1, fg_color="transparent")
        end_frame.pack(side="left", fill="x", expand=True, padx=5)
        ctk.CTkLabel(end_frame, text="End Column:", anchor="w").pack(anchor="w")
        self.end_entry = ctk.CTkEntry(end_frame, placeholder_text="I", height=35)
        self.end_entry.insert(0, "I")
        self.end_entry.pack(fill="x", pady=(5, 0))
        
        check_frame = ctk.CTkFrame(row1, fg_color="transparent")
        check_frame.pack(side="left", fill="x", expand=True, padx=(5, 0))
        ctk.CTkLabel(check_frame, text="Check Column:", anchor="w").pack(anchor="w")
        self.check_entry = ctk.CTkEntry(check_frame, placeholder_text="B", height=35)
        self.check_entry.insert(0, "B")
        self.check_entry.pack(fill="x", pady=(5, 0))
        
        # Row 2: Max row and sentinel settings
        row2 = ctk.CTkFrame(range_inputs, fg_color="transparent")
        row2.pack(fill="x")
        
        maxrow_frame = ctk.CTkFrame(row2, fg_color="transparent")
        maxrow_frame.pack(side="left", fill="x", expand=True, padx=(0, 5))
        ctk.CTkLabel(maxrow_frame, text="Max Row (scan backwards):", anchor="w").pack(anchor="w")
        self.maxrow_entry = ctk.CTkEntry(maxrow_frame, placeholder_text="73", height=35)
        self.maxrow_entry.insert(0, "73")
        self.maxrow_entry.pack(fill="x", pady=(5, 0))
        
        sentinel_frame = ctk.CTkFrame(row2, fg_color="transparent")
        sentinel_frame.pack(side="left", fill="x", expand=True, padx=(5, 0))
        ctk.CTkLabel(sentinel_frame, text="Sentinel Range (change detection):", anchor="w").pack(anchor="w")
        self.sentinel_entry = ctk.CTkEntry(sentinel_frame, placeholder_text="B9:B17", height=35)
        self.sentinel_entry.insert(0, "B9:B17")
        self.sentinel_entry.pack(fill="x", pady=(5, 0))
        
        timeout_frame = ctk.CTkFrame(row2, fg_color="transparent")
        timeout_frame.pack(side="left", fill="x", expand=True, padx=(5, 0))
        ctk.CTkLabel(timeout_frame, text="Timeout (seconds):", anchor="w").pack(anchor="w")
        self.timeout_entry = ctk.CTkEntry(timeout_frame, placeholder_text="10", height=35)
        self.timeout_entry.insert(0, "10")
        self.timeout_entry.pack(fill="x", pady=(5, 0))
        
        # DOWNLOAD SETTINGS
        self.create_section_header("STEP 5: DOWNLOAD SETTINGS")
        
        save_frame = ctk.CTkFrame(self.scrollable, fg_color="transparent")
        save_frame.pack(fill="x", pady=10)
        
        ctk.CTkLabel(save_frame, text="Save Location:", anchor="w").pack(anchor="w", pady=(0, 5))
        
        save_input_frame = ctk.CTkFrame(save_frame, fg_color="transparent")
        save_input_frame.pack(fill="x")
        
        self.save_entry = ctk.CTkEntry(save_input_frame, placeholder_text=DOWNLOADS_DIR, height=35)
        self.save_entry.insert(0, DOWNLOADS_DIR)
        self.save_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        browse_btn = ctk.CTkButton(save_input_frame, text="Browse...", command=self.browse_folder, width=100, height=35)
        browse_btn.pack(side="right")
        
        format_frame = ctk.CTkFrame(self.scrollable, fg_color="transparent")
        format_frame.pack(fill="x", pady=10)
        
        ctk.CTkLabel(format_frame, text="File Format:", anchor="w").pack(anchor="w", pady=(0, 5))
        self.format_dropdown = ctk.CTkComboBox(format_frame, values=["PDF", "Excel (XLSX)", "CSV"], height=35)
        self.format_dropdown.pack(fill="x")
        
        naming_frame = ctk.CTkFrame(self.scrollable, fg_color="transparent")
        naming_frame.pack(fill="x", pady=10)
        
        ctk.CTkLabel(naming_frame, text="File Naming:", anchor="w").pack(anchor="w", pady=(0, 10))
        
        self.naming_var = ctk.StringVar(value="dropdown")
        ctk.CTkRadioButton(naming_frame, text="Use component name", variable=self.naming_var, value="dropdown").pack(anchor="w", pady=2)
        ctk.CTkRadioButton(naming_frame, text="Sequential numbering", variable=self.naming_var, value="sequential").pack(anchor="w", pady=2)
        ctk.CTkRadioButton(naming_frame, text="Timestamp", variable=self.naming_var, value="timestamp").pack(anchor="w", pady=2)
        
        # EXECUTION
        self.create_section_header("STEP 6: START AUTOMATION")
        
        self.start_btn = ctk.CTkButton(
            self.scrollable,
            text="Start Automation",
            command=self.start_automation,
            height=50,
            font=("Arial", 16, "bold"),
            state="disabled"
        )
        self.start_btn.pack(pady=20, fill="x")
        
        self.progress_label = ctk.CTkLabel(self.scrollable, text="Progress: 0/0 (0%)", anchor="w")
        self.progress_label.pack(anchor="w", pady=(10, 5))
        
        self.progress_bar = ctk.CTkProgressBar(self.scrollable)
        self.progress_bar.pack(fill="x", pady=(0, 5))
        self.progress_bar.set(0)
        
        self.current_label = ctk.CTkLabel(self.scrollable, text="", anchor="w", text_color="gray")
        self.current_label.pack(anchor="w")
        
        # LOG
        self.create_section_header("LOG")
        
        self.log_text = ctk.CTkTextbox(self.scrollable, height=150)
        self.log_text.pack(fill="x", pady=10)
        self.log_text.configure(state="disabled")
        
        # Loading overlay (initially hidden)
        self.loading_overlay = None
        
        self.log("Ready. Paste your Google Sheets URL and click 'Connect & Load'")
    
    def create_section_header(self, text):
        header = ctk.CTkLabel(self.scrollable, text=f"--- {text} ---", font=("Arial", 12, "bold"), anchor="w")
        header.pack(anchor="w", pady=(20, 10))
    
    def show_help(self):
        help_text = """
WORKFLOW:
1. Manually set B3 (Menu) to your desired value in Google Sheets
2. Copy and paste the spreadsheet URL here
3. Click 'Connect & Load' - system will:
   - Auto-select 'Extra Component Report' sheet (if exists)
   - Display your current B3 menu selection
   - Load all B6 component values
4. Review the component list
5. Configure data range settings (defaults usually work)
6. Click 'Start Automation'
7. System processes each B6 value:
   - Sets B6 to value
   - Waits for sheet to update (monitors B9:B17)
   - Scans backwards from max row to find data end
   - Exports selected range as PDF/Excel/CSV
   - Moves to next value

TIPS:
- Sentinel Range monitors cells for changes (B9:B17 catches most updates)
- Max Row should be set to your table's maximum possible row
- System scans backwards to find actual data end
        """
        self.log(help_text)
    
    def show_loading_overlay(self, status_text="Loading..."):
        """Show a loading overlay with status indicator"""
        if self.loading_overlay:
            return  # Already showing
        
        # Create semi-transparent overlay
        self.loading_overlay = ctk.CTkFrame(self, fg_color=("gray85", "gray20"))
        self.loading_overlay.place(relx=0, rely=0, relwidth=1, relheight=1)
        
        # Content frame
        content_frame = ctk.CTkFrame(self.loading_overlay, corner_radius=15)
        content_frame.place(relx=0.5, rely=0.5, anchor="center")
        
        # Loading spinner (animated progressbar)
        self.loading_progress = ctk.CTkProgressBar(content_frame, width=300, mode="indeterminate")
        self.loading_progress.pack(padx=40, pady=(40, 20))
        self.loading_progress.start()
        
        # Status label
        self.loading_status_label = ctk.CTkLabel(
            content_frame,
            text=status_text,
            font=("Arial", 14, "bold")
        )
        self.loading_status_label.pack(padx=40, pady=(0, 20))
        
        # Sub-status label for detailed progress
        self.loading_substatus_label = ctk.CTkLabel(
            content_frame,
            text="Please wait...",
            font=("Arial", 11),
            text_color="gray"
        )
        self.loading_substatus_label.pack(padx=40, pady=(0, 40))
        
        # Bring overlay to front
        self.loading_overlay.lift()
    
    def update_loading_status(self, status_text, substatus_text=""):
        """Update the loading overlay status text"""
        if self.loading_overlay and self.loading_status_label:
            self.loading_status_label.configure(text=status_text)
            if substatus_text:
                self.loading_substatus_label.configure(text=substatus_text)
    
    def hide_loading_overlay(self):
        """Hide the loading overlay"""
        if self.loading_overlay:
            if hasattr(self, 'loading_progress'):
                self.loading_progress.stop()
            self.loading_overlay.destroy()
            self.loading_overlay = None
    
    def connect_and_load(self):
        """Connect to spreadsheet and auto-load everything"""
        url = self.url_entry.get().strip()
        
        if not url:
            self.log("Error: Please enter a spreadsheet URL")
            return
        
        self.test_btn.configure(state="disabled", text="Connecting...")
        self.log("Connecting to spreadsheet...")
        
        # Show loading overlay
        self.show_loading_overlay("Connecting to Google Sheets")
        
        def connect_thread():
            # Step 1: Connect to spreadsheet
            self.after(0, lambda: self.update_loading_status(
                "Connecting to Google Sheets",
                "Opening spreadsheet..."
            ))
            success, message = sheets_manager.open_spreadsheet(url)
            
            if success:
                # Step 2: Get worksheets
                self.after(0, lambda: self.update_loading_status(
                    "Connection Successful",
                    "Loading worksheet information..."
                ))
                time.sleep(0.5)  # Brief pause for user to see status
                
                # Step 3: Try to find and load target sheet
                self.after(0, lambda: self.update_loading_status(
                    "Loading Sheet Data",
                    "Searching for 'Extra Component Report'..."
                ))
                time.sleep(0.3)
                
                self.after(0, lambda: self.handle_connection_success(message))
            else:
                self.after(0, lambda: self.handle_connection_failure(message))
        
        threading.Thread(target=connect_thread, daemon=True).start()
    
    def handle_connection_success(self, message):
        """Handle successful connection"""
        self.status_label.configure(text=f"Status: Connected - {message}", text_color="#2fa572")
        self.log(f"Success: {message}")
        
        # Get all worksheets
        worksheets = sheets_manager.get_worksheet_names()
        self.log(f"Found {len(worksheets)} sheet(s)")
        
        # Update loading status
        self.update_loading_status(
            "Loading Sheet Data",
            f"Found {len(worksheets)} worksheets"
        )
        
        # Try to auto-select "Extra Component Report"
        target_sheet = "Extra Component Report"
        if target_sheet in worksheets:
            self.log(f"Auto-selecting sheet: {target_sheet}")
            self.update_loading_status(
                "Loading Components",
                f"Reading data from '{target_sheet}'..."
            )
            self.load_sheet_data(target_sheet)
        else:
            self.log(f"'{target_sheet}' not found. Available sheets: {', '.join(worksheets)}")
            self.log("Please manually select the correct sheet in Google Sheets")
            self.refresh_btn.configure(state="normal")
            self.hide_loading_overlay()
        
        self.test_btn.configure(state="normal", text="Connected ✓")
    
    def handle_connection_failure(self, message):
        """Handle connection failure"""
        self.hide_loading_overlay()
        self.status_label.configure(text=f"Status: Error - {message}", text_color="#f56c6c")
        self.log(f"Error: {message}")
        self.test_btn.configure(state="normal", text="Connect & Load")
    
    def load_sheet_data(self, sheet_name):
        """Load data from the selected sheet"""
        self.log(f"Loading data from '{sheet_name}'...")
        
        def load_thread():
            # Update status: Getting worksheet
            self.after(0, lambda: self.update_loading_status(
                "Loading Components",
                "Accessing worksheet..."
            ))
            
            worksheet = sheets_manager.get_worksheet(sheet_name)
            if not worksheet:
                self.after(0, lambda: self.log(f"Error: Could not access sheet '{sheet_name}'"))
                self.after(0, self.hide_loading_overlay)
                return
            
            # Update status: Reading menu value
            self.after(0, lambda: self.update_loading_status(
                "Loading Components",
                "Reading menu value (B3)..."
            ))
            time.sleep(0.3)
            menu_value = sheets_manager.get_cell_value(worksheet, self.menu_display_cell)
            
            # Update status: Reading dropdown values
            self.after(0, lambda: self.update_loading_status(
                "Loading Components",
                "Reading component dropdown (B6)..."
            ))
            time.sleep(0.3)
            component_values = sheets_manager.read_dropdown_values_from_cell(
                worksheet, 
                self.component_dropdown_cell,
                sheet_name
            )
            
            # Update status: Processing complete
            self.after(0, lambda: self.update_loading_status(
                "Complete",
                f"Loaded {len(component_values)} components"
            ))
            time.sleep(0.5)
            
            self.after(0, lambda: self.handle_loaded_data(menu_value, component_values))
            self.after(0, self.hide_loading_overlay)
        
        threading.Thread(target=load_thread, daemon=True).start()
    
    def handle_loaded_data(self, menu_value, component_values):
        """Handle loaded data and update UI"""
        # Display menu value
        if menu_value:
            self.current_menu_value = menu_value
            self.menu_display.configure(state="normal")
            self.menu_display.delete(0, "end")
            self.menu_display.insert(0, menu_value)
            self.menu_display.configure(state="disabled")
            self.log(f"Current Menu (B3): {menu_value}")
        else:
            self.log("Warning: B3 appears empty. Please set your menu value in Google Sheets")
        
        # Display component values
        if component_values:
            self.component_values = component_values
            self.log(f"Loaded {len(component_values)} component(s) from B6")
            
            self.component_preview_label.configure(
                text=f"Found {len(component_values)} components to process",
                text_color="#2fa572"
            )
            
            self.component_preview_box.delete("1.0", "end")
            for i, comp in enumerate(component_values, 1):
                self.component_preview_box.insert("end", f"{i}. {comp}\n")
            
            self.start_btn.configure(state="normal")
            self.refresh_btn.configure(state="normal")
            self.log("Ready! Click 'Start Automation' when ready")
        else:
            self.log("Error: No component values found in B6 dropdown")
            self.component_preview_label.configure(
                text="No components found. Check B3 value and click 'Refresh B6 Values'",
                text_color="#f56c6c"
            )
            self.refresh_btn.configure(state="normal")
    
    def refresh_component_values(self):
        """Refresh B6 values after user changes B3"""
        self.log("Refreshing component values...")
        self.refresh_btn.configure(state="disabled", text="Refreshing...")
        
        # Show loading overlay for refresh
        self.show_loading_overlay("Refreshing Components")
        
        sheet_name = "Extra Component Report"
        self.load_sheet_data(sheet_name)
        
        self.refresh_btn.configure(state="normal", text="Refresh B6 Values")
    
    def start_automation(self):
        if self.is_running:
            self.log("Automation already running")
            return
        
        if not self.component_values:
            self.log("Error: No components to process")
            return
        
        # Reset failed components list
        self.failed_components = []
        
        self.start_btn.configure(state="disabled", text="Running...", fg_color="#e6a23c")
        self.is_running = True
        
        self.log("=" * 40)
        self.log("Starting automation...")
        self.log(f"Menu: {self.current_menu_value}")
        self.log(f"Processing {len(self.component_values)} components")
        self.log("=" * 40)
        
        thread = threading.Thread(target=self.run_automation, daemon=True)
        thread.start()
    
    def run_automation(self):
            try:
                sheet_name = "Extra Component Report"
                dropdown_cell = self.component_dropdown_cell
                start_cell = self.start_entry.get().strip()
                end_column = self.end_entry.get().strip().upper()
                check_column = self.check_entry.get().strip().upper()
                max_row = int(self.maxrow_entry.get().strip())
                sentinel_range = self.sentinel_entry.get().strip()
                timeout = int(self.timeout_entry.get().strip())
                save_location = self.save_entry.get().strip()
                file_format = self.format_dropdown.get()
                naming_mode = self.naming_var.get()
                
                if not os.path.exists(save_location):
                    os.makedirs(save_location, exist_ok=True)
                
                start_row = int(''.join(filter(str.isdigit, start_cell)))
                start_col = ''.join(filter(str.isalpha, start_cell))
                
                worksheet = sheets_manager.get_worksheet(sheet_name)
                if not worksheet:
                    self.log("Error: Could not access worksheet")
                    return
                
                original_value = sheets_manager.get_cell_value(worksheet, dropdown_cell)
                
                total = len(self.component_values)
                success_count = 0
                failed_count = 0
                
                for idx, value in enumerate(self.component_values, 1):
                    if not self.is_running:
                        self.log("Automation stopped by user")
                        break
                    
                    progress = idx / total
                    self.update_progress(idx, total, progress, f"Processing: {value}")
                    
                    self.log(f"[{idx}/{total}] Processing: '{value}'")
                    
                    try:
                        # Read sentinel values BEFORE setting B6
                        self.log(f"  Reading sentinel range: {sentinel_range}")
                        initial_sentinel = worksheet.get(sentinel_range)
                        
                        # Set B6 to new value
                        self.log(f"  Setting {dropdown_cell} to: {value}")
                        if not sheets_manager.set_cell_value(worksheet, dropdown_cell, value):
                            self.log("  Error: Could not set dropdown value")
                            self.failed_components.append({
                                'name': value,
                                'reason': 'Could not set dropdown value'
                            })
                            failed_count += 1
                            continue
                        
                        # Wait for sheet to update (monitor sentinel)
                        self.log(f"  Waiting for sheet update (monitoring {sentinel_range})...")
                        change_detected = self.wait_for_change(
                            worksheet, 
                            sentinel_range, 
                            initial_sentinel, 
                            timeout
                        )
                        
                        if change_detected:
                            self.log("  Sheet updated successfully")
                        else:
                            self.log(f"  Warning: No change detected after {timeout}s, proceeding anyway")
                        
                        # Find last row by scanning backwards from max_row
                        self.log(f"  Scanning backwards from row {max_row}...")
                        last_row = self.find_last_row_backwards(worksheet, check_column, start_row, max_row)
                        self.log(f"  Data ends at row: {last_row}")
                        
                        data_range = f"{start_col}{start_row}:{end_column}{last_row}"
                        self.log(f"  Exporting range: {data_range}")
                        
                        filename = self.generate_filename(value, idx, naming_mode, file_format)
                        output_path = os.path.join(save_location, filename)
                        
                        if file_format == "PDF":
                            success, msg = sheets_manager.export_range_as_pdf(sheet_name, data_range, output_path)
                        elif file_format == "Excel (XLSX)":
                            success, msg = sheets_manager.export_range_as_excel(sheet_name, data_range, output_path)
                        elif file_format == "CSV":
                            success, msg = sheets_manager.export_range_as_csv(sheet_name, data_range, output_path)
                        else:
                            success, msg = False, "Unknown format"
                        
                        if success:
                            self.log(f"  ✓ Saved: {filename}")
                            success_count += 1
                        else:
                            self.log(f"  ✗ Failed: {msg}")
                            self.failed_components.append({
                                'name': value,
                                'reason': msg
                            })
                            failed_count += 1
                        
                    except Exception as e:
                        error_msg = str(e)
                        self.log(f"  ✗ Error: {error_msg}")
                        self.failed_components.append({
                            'name': value,
                            'reason': error_msg
                        })
                        failed_count += 1
                
                # Restore original value
                if original_value:
                    self.log(f"Restoring original {dropdown_cell} value: {original_value}")
                    sheets_manager.set_cell_value(worksheet, dropdown_cell, original_value)
                
                self.log("=" * 40)
                self.log(f"COMPLETE! Success: {success_count}, Failed: {failed_count}")
                self.log("=" * 40)
                
                self.show_completion_dialog(success_count, failed_count, save_location)
                
            except Exception as e:
                self.log(f"Critical error: {str(e)}")
                import traceback
                self.log(traceback.format_exc())
            
            finally:
                self.is_running = False
                self.after(0, lambda: self.start_btn.configure(state="normal", text="Start Automation", fg_color=["#3B8ED0", "#1F6AA5"]))
                self.after(0, lambda: self.update_progress(0, 0, 0, ""))
    
    def wait_for_change(self, worksheet, sentinel_range, initial_values, timeout):
        """Wait for sentinel range to change (indicates sheet has updated)"""
        start_time = time.time()
        check_interval = 0.5  # Check every 0.5 seconds
        
        while time.time() - start_time < timeout:
            try:
                current_values = worksheet.get(sentinel_range)
                if current_values != initial_values:
                    return True  # Change detected
            except Exception as e:
                self.log(f"  Warning: Error reading sentinel: {e}")
            
            time.sleep(check_interval)
        
        return False  # Timeout reached
    
    def find_last_row_backwards(self, worksheet, column_letter, start_row, max_row):
        """Scan backwards from max_row to find last row with data"""
        try:
            col_num = sheets_manager._col_letter_to_num(column_letter)
            column_values = worksheet.col_values(col_num)
            
            # Scan backwards from max_row
            for row_num in range(min(max_row, len(column_values)), start_row - 1, -1):
                if row_num - 1 < len(column_values):
                    cell_value = column_values[row_num - 1]
                    if cell_value and str(cell_value).strip():
                        return row_num
            
            # If no data found, return start_row
            return start_row
            
        except Exception as e:
            self.log(f"  Error in backward scan: {e}")
            return start_row
    
    def generate_filename(self, dropdown_value, index, naming_mode, file_format):
        clean_value = "".join(c for c in dropdown_value if c.isalnum() or c in (' ', '_', '-')).strip()
        clean_value = clean_value.replace(' ', '_')
        
        ext_map = {"PDF": "pdf", "Excel (XLSX)": "xlsx", "CSV": "csv"}
        ext = ext_map.get(file_format, "pdf")
        
        if naming_mode == "dropdown":
            return f"{clean_value}.{ext}"
        elif naming_mode == "sequential":
            return f"Report_{index}.{ext}"
        else:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            return f"{timestamp}_{index}.{ext}"
    
    def show_completion_dialog(self, success, failed, location):
        def show():
            dialog = ctk.CTkToplevel(self)
            dialog.title("Automation Complete")
            dialog.geometry("600x500")
            dialog.transient(self)
            dialog.grab_set()
            
            # Center the dialog
            dialog.update_idletasks()
            x = (dialog.winfo_screenwidth() // 2) - (300)
            y = (dialog.winfo_screenheight() // 2) - (250)
            dialog.geometry(f"600x500+{x}+{y}")
            
            # Main container with padding
            main_frame = ctk.CTkFrame(dialog)
            main_frame.pack(fill="both", expand=True, padx=20, pady=20)
            
            # Title
            title_label = ctk.CTkLabel(
                main_frame,
                text="Automation Complete",
                font=("Arial", 18, "bold")
            )
            title_label.pack(pady=(0, 15))
            
            # Summary stats
            stats_frame = ctk.CTkFrame(main_frame)
            stats_frame.pack(fill="x", pady=(0, 15))
            
            success_label = ctk.CTkLabel(
                stats_frame,
                text=f"✓ Successful: {success}",
                font=("Arial", 13),
                text_color="#2fa572"
            )
            success_label.pack(pady=5)
            
            failed_label = ctk.CTkLabel(
                stats_frame,
                text=f"✗ Failed: {failed}",
                font=("Arial", 13),
                text_color="#f56c6c" if failed > 0 else "gray"
            )
            failed_label.pack(pady=5)
            
            # Failed components section (only show if there are failures)
            if failed > 0 and self.failed_components:
                # Warning message
                warning_frame = ctk.CTkFrame(main_frame, fg_color="#e6a23c")
                warning_frame.pack(fill="x", pady=(10, 10))
                
                warning_label = ctk.CTkLabel(
                    warning_frame,
                    text="⚠ Failed Components: Kindly double-check the actual contents on the sheet. Thank you.",
                    font=("Arial", 11, "bold"),
                    text_color="white",
                    wraplength=540
                )
                warning_label.pack(pady=10, padx=10)
                
                # Failed components list label
                failed_list_label = ctk.CTkLabel(
                    main_frame,
                    text="Failed Components:",
                    font=("Arial", 12, "bold"),
                    anchor="w"
                )
                failed_list_label.pack(anchor="w", pady=(5, 5))
                
                # Scrollable text box for failed components
                failed_textbox = ctk.CTkTextbox(
                    main_frame,
                    height=200,
                    wrap="word"
                )
                failed_textbox.pack(fill="both", expand=True, pady=(0, 10))
                
                # Populate failed components list
                for i, failed_item in enumerate(self.failed_components, 1):
                    component_name = failed_item['name']
                    reason = failed_item['reason']
                    failed_textbox.insert("end", f"{i}. {component_name}\n")
                    failed_textbox.insert("end", f"   Reason: {reason}\n\n")
                
                failed_textbox.configure(state="disabled")
            
            # Button frame
            button_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
            button_frame.pack(pady=(10, 0))
            
            def open_folder():
                try:
                    if os.name == 'nt':  # Windows
                        os.startfile(location)
                    elif os.name == 'posix':  # macOS and Linux
                        import subprocess
                        if os.uname().sysname == 'Darwin':  # macOS
                            subprocess.run(['open', location])
                        else:  # Linux
                            subprocess.run(['xdg-open', location])
                except Exception as e:
                    self.log(f"Could not open folder: {location} - {str(e)}")
            
            open_btn = ctk.CTkButton(
                button_frame,
                text="Open Downloads Folder",
                command=open_folder,
                width=180,
                height=35
            )
            open_btn.pack(side="left", padx=5)
            
            close_btn = ctk.CTkButton(
                button_frame,
                text="Close",
                command=dialog.destroy,
                width=100,
                height=35
            )
            close_btn.pack(side="left", padx=5)
        
        self.after(0, show)
    
    def update_progress(self, current, total, value, text):
        def update():
            if total > 0:
                self.progress_label.configure(text=f"Progress: {current}/{total} ({int(value*100)}%)")
                self.progress_bar.set(value)
            else:
                self.progress_label.configure(text="Progress: 0/0 (0%)")
                self.progress_bar.set(0)
            self.current_label.configure(text=text)
        self.after(0, update)
    
    def browse_folder(self):
        try:
            folder = filedialog.askdirectory(initialdir=self.save_entry.get() or DOWNLOADS_DIR)
            if folder:
                self.save_entry.delete(0, "end")
                self.save_entry.insert(0, folder)
                self.log(f"Save location: {folder}")
        except Exception as e:
            self.log(f"Error: {str(e)}")
    
    def log(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        def update_log():
            self.log_text.configure(state="normal")
            self.log_text.insert("end", f"[{timestamp}] {message}\n")
            self.log_text.see("end")
            self.log_text.configure(state="disabled")
        self.after(0, update_log)
    
    def go_back(self):
        if self.is_running:
            self.log("Cannot go back while automation is running")
            return
        self.on_back()
    
    def stop_automation(self):
        self.is_running = False
        self.log("Stopping automation...")