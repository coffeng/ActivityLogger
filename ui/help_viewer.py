"""
Help viewer window
"""
import tkinter as tk
from tkinter import scrolledtext


class HelpViewer:
    """Help dialog window"""
    
    def __init__(self, log_path):
        self.root = tk.Tk()
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
            command=self.root.destroy,
            width=10
        )
        ok_button.pack(pady=10)
        
        # Center the window
        self.root.transient()
        self.root.grab_set()
        
        self.root.mainloop()