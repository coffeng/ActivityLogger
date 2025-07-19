"""
Help viewer window
"""
import tkinter as tk
from tkinter import scrolledtext


class HelpViewer:
    """Help dialog window"""
    
    def __init__(self, parent_or_log_path, log_path=None):
        # Handle both cases: HelpViewer(parent, log_path) and HelpViewer(log_path)
        if log_path is None:
            # Called with just log_path (from tray)
            self.root = tk.Tk()
            log_path = parent_or_log_path
            self.has_parent = False
        else:
            # Called with parent and log_path (from viewer)
            self.root = tk.Toplevel(parent_or_log_path)
            self.has_parent = True
            self.parent = parent_or_log_path
            
        self.root.title("Activity Logger Help")
        self.root.geometry("600x500")
        
        # Help text
        help_text = f"""
Activity Logger - Help

MENU ITEMS:

Open Log File: Opens the activity log viewer window
Start Logging: Begins monitoring active windows and applications
Stop Logging: Pauses activity monitoring (keeps existing data)
Restart Logging: Stops and restarts the logging process
Help: Shows this help window
Exit: Closes the application completely

FEATURES:

• Automatic window tracking - logs when you switch between applications
• Idle detection - detects when you're away from computer
• Category assignment - automatically categorizes applications
• Real-time viewing - see your activity as it's logged
• Statistics - shows session duration, total logged time, and idle time

LOG LOCATION:
{log_path}

USAGE TIPS:

• The application runs in the system tray (bottom-right corner)
• Right-click the tray icon to access menu options
• The log file is a CSV that can be opened in Excel or other tools
• Categories can be customized by right-clicking in the Summary tab
• Idle detection varies by application type (longer for meetings)

TROUBLESHOOTING:

• If logging stops working, try "Restart Logging"
• Log files are saved automatically and safely
• Multiple log viewer windows are prevented automatically
• The application starts logging automatically when launched

For more information or support, check the source code comments.
        """
        
        # Create scrolled text widget
        text_widget = scrolledtext.ScrolledText(
            self.root, 
            wrap=tk.WORD, 
            font=('Arial', 10),
            padx=10,
            pady=10
        )
        text_widget.pack(fill=tk.BOTH, expand=True)
        text_widget.insert(tk.END, help_text)
        text_widget.config(state=tk.DISABLED)  # Make read-only
        
        # OK button
        ok_button = tk.Button(
            self.root, 
            text="OK", 
            command=self.on_close,
            width=10
        )
        ok_button.pack(pady=10)
        
        # Center the window
        if self.has_parent:
            self.root.transient(self.parent)
            self.root.grab_set()

        # Handle window close from 'X' button
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        
        # Only call mainloop if no parent (called from tray)
        if not self.has_parent:
            self.root.mainloop()

    def on_close(self):
        """Handle window close event"""
        try:
            # Force withdraw the window first
            self.root.withdraw()
            
            # Always try to quit mainloop (like OK button does)
            try:
                self.root.quit()
            except Exception:
                pass
            
            # Destroy the window and all widgets
            self.root.destroy()
            
        except Exception as e:
            print(f"Error in help_viewer on_close: {e}")
            try:
                if hasattr(self, 'root') and self.root:
                    self.root.quit()
                    self.root.destroy()
            except Exception:
                pass