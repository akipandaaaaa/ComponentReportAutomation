"""
Main hub screen with automation selection grid
Auto-connects to Google Sheets on startup
"""
import customtkinter as ctk
from utils.google_sheets import sheets_manager

class MainHubScreen(ctk.CTkFrame):
    def __init__(self, parent, on_automation_select):
        super().__init__(parent)
        self.on_automation_select = on_automation_select
        
        # Title
        title = ctk.CTkLabel(
            self,
            text="Automation Hub",
            font=("Arial", 28, "bold")
        )
        title.pack(pady=30)
        
        # Connection status
        self.status_frame = ctk.CTkFrame(self)
        self.status_frame.pack(pady=10, padx=40, fill="x")
        
        self.status_label = ctk.CTkLabel(
            self.status_frame,
            text="● Connecting...",
            font=("Arial", 12),
            text_color="orange"
        )
        self.status_label.pack(side="left", padx=10, pady=5)
        
        self.connect_btn = ctk.CTkButton(
            self.status_frame,
            text="Connecting...",
            command=self.connect_google_sheets,
            width=180,
            state="disabled"
        )
        self.connect_btn.pack(side="right", padx=10, pady=5)
        
        # Grid container
        grid_frame = ctk.CTkFrame(self)
        grid_frame.pack(pady=20, padx=40, fill="both", expand=True)
        
        # Grid label
        grid_label = ctk.CTkLabel(
            grid_frame,
            text="Select Automation:",
            font=("Arial", 16)
        )
        grid_label.pack(pady=(10, 20))
        
        # Create 2x2 grid
        grid_container = ctk.CTkFrame(grid_frame, fg_color="transparent")
        grid_container.pack(expand=True)
        
        # Configure grid
        for i in range(2):
            grid_container.grid_columnconfigure(i, weight=1, pad=15)
            grid_container.grid_rowconfigure(i, weight=1, pad=15)
        
        # Button 1: Component Report Download
        self.create_grid_button(
            grid_container,
            "Component Report\nDownload",
            "Download reports from\nGoogle Sheets dropdowns",
            0, 0,
            self.open_component_report
        )
        
        # Button 2-4: Placeholder
        for i in range(1, 4):
            row = i // 2
            col = i % 2
            self.create_grid_button(
                grid_container,
                f"Future\nAutomation {i+1}",
                "Coming soon...",
                row, col,
                None,
                enabled=False
            )
        
        # Auto-connect on startup
        self.after(500, self.connect_google_sheets)
    
    def create_grid_button(self, parent, title, description, row, col, command, enabled=True):
        """Create a styled grid button"""
        btn_frame = ctk.CTkFrame(
            parent,
            width=200,
            height=150,
            corner_radius=10
        )
        btn_frame.grid(row=row, column=col, padx=10, pady=10, sticky="nsew")
        btn_frame.grid_propagate(False)
        
        # Title
        title_label = ctk.CTkLabel(
            btn_frame,
            text=title,
            font=("Arial", 16, "bold"),
            wraplength=180
        )
        title_label.pack(pady=(20, 5))
        
        # Description
        desc_label = ctk.CTkLabel(
            btn_frame,
            text=description,
            font=("Arial", 11),
            text_color="gray",
            wraplength=180
        )
        desc_label.pack(pady=5)
        
        # Action button
        if enabled:
            action_btn = ctk.CTkButton(
                btn_frame,
                text="Open",
                command=command,
                width=100,
                height=32
            )
            action_btn.pack(pady=(10, 10))
        else:
            action_btn = ctk.CTkButton(
                btn_frame,
                text="Coming Soon",
                command=None,
                width=100,
                height=32,
                state="disabled"
            )
            action_btn.pack(pady=(10, 10))
    
    def connect_google_sheets(self):
        """Connect to Google Sheets API"""
        self.connect_btn.configure(state="disabled", text="Connecting...")
        self.status_label.configure(text="● Connecting...", text_color="orange")
        
        success, message = sheets_manager.connect()
        
        if success:
            self.status_label.configure(
                text="● Connected",
                text_color="#2fa572"
            )
            self.connect_btn.configure(text="Connected ✓", state="disabled")
        else:
            self.status_label.configure(
                text="● Connection Failed",
                text_color="#f56c6c"
            )
            self.connect_btn.configure(text="Retry Connection", state="normal")
            
            # Show error dialog with details
            self.show_connection_error(message)
    
    def show_connection_error(self, message):
        """Show detailed connection error dialog"""
        error_dialog = ctk.CTkToplevel(self)
        error_dialog.title("Connection Error")
        error_dialog.geometry("500x300")
        error_dialog.transient(self)
        error_dialog.grab_set()
        
        # Center the dialog
        error_dialog.update_idletasks()
        x = (error_dialog.winfo_screenwidth() // 2) - (500 // 2)
        y = (error_dialog.winfo_screenheight() // 2) - (300 // 2)
        error_dialog.geometry(f"500x300+{x}+{y}")
        
        error_label = ctk.CTkLabel(
            error_dialog,
            text="⚠ Connection Failed",
            font=("Arial", 18, "bold"),
            text_color="#f56c6c"
        )
        error_label.pack(pady=20)
        
        error_msg = ctk.CTkTextbox(
            error_dialog,
            height=120,
            wrap="word"
        )
        error_msg.pack(pady=10, padx=20, fill="both")
        error_msg.insert("1.0", f"Error: {message}\n\n")
        error_msg.insert("end", "Common issues:\n")
        error_msg.insert("end", "1. credentials.json file is missing\n")
        error_msg.insert("end", "2. Service account not set up properly\n")
        error_msg.insert("end", "3. Google Sheets API not enabled\n\n")
        error_msg.insert("end", "Please check your setup and try again.")
        error_msg.configure(state="disabled")
        
        close_btn = ctk.CTkButton(
            error_dialog,
            text="Close",
            command=error_dialog.destroy,
            width=120
        )
        close_btn.pack(pady=20)
    
    def open_component_report(self):
        """Open Component Report automation"""
        if not sheets_manager.connected:
            # Show warning
            warning = ctk.CTkToplevel(self)
            warning.title("Not Connected")
            warning.geometry("350x180")
            warning.transient(self)
            warning.grab_set()
            
            # Center the dialog
            warning.update_idletasks()
            x = (warning.winfo_screenwidth() // 2) - (350 // 2)
            y = (warning.winfo_screenheight() // 2) - (180 // 2)
            warning.geometry(f"350x180+{x}+{y}")
            
            label = ctk.CTkLabel(
                warning,
                text="⚠ Not Connected to Google Sheets",
                font=("Arial", 14, "bold")
            )
            label.pack(pady=20)
            
            msg = ctk.CTkLabel(
                warning,
                text="Please connect to Google Sheets first\nbefore opening this automation.",
                wraplength=300
            )
            msg.pack(pady=10)
            
            btn_frame = ctk.CTkFrame(warning, fg_color="transparent")
            btn_frame.pack(pady=20)
            
            retry_btn = ctk.CTkButton(
                btn_frame,
                text="Retry Connection",
                command=lambda: [warning.destroy(), self.connect_google_sheets()]
            )
            retry_btn.pack(side="left", padx=5)
            
            cancel_btn = ctk.CTkButton(
                btn_frame,
                text="Cancel",
                command=warning.destroy
            )
            cancel_btn.pack(side="left", padx=5)
            return
        
        self.on_automation_select("component_report")