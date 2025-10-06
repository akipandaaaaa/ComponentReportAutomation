"""
Main application entry point
Automation Hub - Component Report Download
"""
import customtkinter as ctk
from screens.main_hub import MainHubScreen
from screens.component_report import ComponentReportScreen
from utils.config import APP_NAME, WINDOW_SIZE

class AutomationApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Configure window
        self.title(APP_NAME)
        self.geometry(WINDOW_SIZE)
        
        # Set theme
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        
        # Container for screens
        self.container = ctk.CTkFrame(self)
        self.container.pack(fill="both", expand=True)
        
        # Initialize screens
        self.screens = {}
        self.current_screen = None
        
        # Create main hub
        self.show_screen("main_hub")
    
    def show_screen(self, screen_name):
        """Show a specific screen"""
        # Hide current screen
        if self.current_screen:
            self.current_screen.pack_forget()
        
        # Create screen if it doesn't exist
        if screen_name not in self.screens:
            if screen_name == "main_hub":
                self.screens[screen_name] = MainHubScreen(
                    self.container,
                    on_automation_select=self.show_screen
                )
            elif screen_name == "component_report":
                self.screens[screen_name] = ComponentReportScreen(
                    self.container,
                    on_back=lambda: self.show_screen("main_hub")
                )
        
        # Show the screen
        self.current_screen = self.screens[screen_name]
        self.current_screen.pack(fill="both", expand=True)

if __name__ == "__main__":
    app = AutomationApp()
    app.mainloop()