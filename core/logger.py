"""
Main activity logging functionality
"""
import os
import csv
import time
import threading
import datetime
import psutil
import win32gui
import win32process
import win32con
import win32api
from ctypes import wintypes
import ctypes
from ctypes import WINFUNCTYPE

from .config import ConfigManager
from .utils import get_log_path

# Constants
LOG_PATH = get_log_path()
INTERVAL_SECONDS = 1

# Windows API setup
user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32
HOOKPROC = WINFUNCTYPE(ctypes.c_int, ctypes.c_int, wintypes.WPARAM, wintypes.LPARAM)


class ActivityLogger:
    """Main activity logger class"""
    
    def __init__(self, log_path=LOG_PATH, interval=INTERVAL_SECONDS):
        self.log_path = log_path
        self.interval = interval
        self.running = False
        self.thread = None
        self.prev_window = ""
        self.prev_details = ""
        self.prev_category = ""
        self.start_time = None
        self.was_idle = False
        self.hook = None
        self.hook_id = None

        # Track application start time for statistics
        self.app_start_time = datetime.datetime.now()
        
        # Initialize configuration manager
        self.config_manager = ConfigManager(self.log_path)
        self.app_categories = self.config_manager.load_app_categories()
        
        # Track rows added since last config save
        self.rows_added_since_save = 0
        self.save_interval = 10  # Save config every 10 new rows

        # For idle detection
        self.idle_check_interval = 5  # Check idle every 5 seconds

    def save_app_categories(self):
        """Save app categories using config manager"""
        self.config_manager.save_app_categories(self.app_categories)

    def get_active_window_title(self):
        """Get the title of the currently active window"""
        hwnd = win32gui.GetForegroundWindow()
        if hwnd == 0:
            return ""
        return win32gui.GetWindowText(hwnd)

    def get_active_process_name(self):
        """Get the name of the process for the currently active window"""
        hwnd = win32gui.GetForegroundWindow()
        if hwnd == 0:
            return ""
        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            return psutil.Process(pid).name()
        except Exception:
            return ""

    def get_window_details(self, window_title, process_name):
        """Extract meaningful details from window title"""
        try:
            proc = process_name.lower()
            title = window_title
            # Remove process name or known suffixes from title
            if proc == "excel.exe" and " - Excel" in title:
                return title.replace(" - Excel", "").strip()
            elif proc == "winword.exe" and " - Word" in title:
                return title.replace(" - Word", "").strip()
            elif proc == "powerpnt.exe" and " - PowerPoint" in title:
                return title.replace(" - PowerPoint", "").strip()
            elif proc == "outlook.exe" and " - Outlook" in title:
                return title.replace(" - Outlook", "").strip()
            elif proc == "chrome.exe" and " - Google Chrome" in title:
                return title.replace(" - Google Chrome", "").strip()
            elif proc == "msedge.exe" and " - Microsoft Edge" in title:
                return title.replace(" - Microsoft Edge", "").strip()
            elif " - " in title:
                return title.rsplit(" - ", 1)[0].strip()
            else:
                return title.strip()
        except Exception:
            pass
        return window_title

    def get_category(self, window_title, process_name, window_details):
        """Get category using loaded app_categories"""
        proc = process_name.lower()
        title = window_title.lower()
        details = window_details.lower()
        
        for key, cat in self.app_categories.items():
            if key in proc or key in title or key in details:
                return cat
        return "Uncategorized"

    def get_idle_seconds(self):
        """Get the number of seconds since last user input"""
        class LASTINPUTINFO(ctypes.Structure):
            _fields_ = [("cbSize", ctypes.c_uint), ("dwTime", ctypes.c_uint)]
        lii = LASTINPUTINFO()
        lii.cbSize = ctypes.sizeof(LASTINPUTINFO)
        if ctypes.windll.user32.GetLastInputInfo(ctypes.byref(lii)):
            millis = win32api.GetTickCount() - lii.dwTime
            return millis / 1000.0
        return 0

    def log_activity(self, start, end, window, proc, details, category):
        """Log an activity to the CSV file"""
        duration = int((end - start).total_seconds())
        file_exists = os.path.isfile(self.log_path)
        with open(self.log_path, "a", newline='', encoding="utf-8") as f:
            writer = csv.writer(f)
            if not file_exists:
                writer.writerow([
                    "StartTime", "EndTime", "DurationSeconds", "WindowTitle", 
                    "WindowDetails", "ProcessName", "Category"
                ])
            writer.writerow([
                start.strftime("%Y-%m-%d %H:%M:%S"),
                end.strftime("%Y-%m-%d %H:%M:%S"),
                duration,
                window,
                details,
                proc,
                category
            ])
        
        # Increment row counter and save config if needed
        self.rows_added_since_save += 1
        if self.rows_added_since_save >= self.save_interval:
            self.save_app_categories()
            self.rows_added_since_save = 0

    def install_hook(self):
        """Install the window change hook"""
        try:
            def win_event_proc(hWinEventHook, event, hwnd, idObject, idChild, dwEventThread, dwmsEventTime):
                try:
                    if event == 0x0003:  # EVENT_SYSTEM_FOREGROUND
                        self.process_window_change()
                except Exception as e:
                    print(f"Hook callback error: {e}")

            from ctypes import WINFUNCTYPE
            WinEventProc = WINFUNCTYPE(None, wintypes.HANDLE, wintypes.DWORD,
                                       wintypes.HWND, wintypes.LONG, wintypes.LONG,
                                       wintypes.DWORD, wintypes.DWORD)

            self.hook = WinEventProc(win_event_proc)

            self.hook_id = user32.SetWinEventHook(
                0x0003, 0x0003, None, self.hook, 0, 0, 0x0000
            )

            if self.hook_id:
                print("WinEvent hook installed successfully")
                return True
            else:
                print("SetWinEventHook returned NULL")
                return False

        except Exception as e:
            print(f"Failed to install hook: {e}")
            return False
        
    def uninstall_hook(self):
        """Uninstall the window change hook"""
        try:
            if hasattr(self, 'hook_id') and self.hook_id:
                result = user32.UnhookWinEvent(self.hook_id)
                if result:
                    print("WinEvent hook uninstalled successfully")
                else:
                    print("Failed to uninstall WinEvent hook")
                self.hook_id = None
            
            if hasattr(self, 'hook'):
                self.hook = None
                
        except Exception as e:
            print(f"Error uninstalling hook: {e}")
  
    def process_window_change(self):
        """Process the window change event"""
        try:
            current_window = self.get_active_window_title()
            current_proc = self.get_active_process_name()
            current_details = self.get_window_details(current_window, current_proc)
            current_category = self.get_category(current_window, current_proc, current_details)

            # Check if this is actually a change we care about
            if (current_window != self.prev_window or
                current_details != self.prev_details or
                    current_category != self.prev_category):

                # Log the previous window if we have one
                if self.prev_window and self.start_time and not self.was_idle:
                    end_time = datetime.datetime.now()
                    self.log_activity(
                        self.start_time, end_time,
                        self.prev_window, getattr(self, 'prev_proc', ''),
                        self.prev_details, self.prev_category
                    )

                # Update to new window
                self.prev_window = current_window
                self.prev_proc = current_proc
                self.prev_details = current_details
                self.prev_category = current_category
                self.start_time = datetime.datetime.now()

        except Exception as e:
            print(f"Error processing window change: {e}")

    def logger_loop(self):
        """Main logger loop - uses polling for reliability"""
        print("Using polling method for reliability")
        self.polling_fallback()

    def polling_fallback(self):
        """Enhanced polling method"""
        print("Starting polling method")

        self.start_time = datetime.datetime.now()
        self.prev_window = self.get_active_window_title()
        self.prev_proc = self.get_active_process_name()
        self.prev_details = self.get_window_details(self.prev_window, self.prev_proc)
        self.prev_category = self.get_category(self.prev_window, self.prev_proc, self.prev_details)
        self.was_idle = False
        self.idle_start = None

        # Polling intervals
        window_check_interval = 0.5  # Check every 500ms for window changes
        idle_check_counter = 0
        idle_check_frequency = 10  # Check idle every 5 seconds (10 * 0.5s)

        print(f"Initial window: {self.prev_window}")

        while self.running:
            try:
                # Get current window info
                current_window = self.get_active_window_title()
                current_proc = self.get_active_process_name()
                current_details = self.get_window_details(current_window, current_proc)
                current_category = self.get_category(current_window, current_proc, current_details)

                # Check if window changed
                if (current_window != self.prev_window or
                    current_details != self.prev_details or
                    current_category != self.prev_category):

                    print(f"Window changed: {self.prev_window} -> {current_window}")

                    # Log previous activity if we have one
                    if self.prev_window and self.start_time and not self.was_idle:
                        end_time = datetime.datetime.now()
                        duration = (end_time - self.start_time).total_seconds()
                        if duration >= 1:  # Only log if activity lasted at least 1 second
                            self.log_activity(
                                self.start_time, end_time,
                                self.prev_window, self.prev_proc,
                                self.prev_details, self.prev_category
                            )
                            print(f"Logged activity: {self.prev_window} ({duration:.1f}s)")

                    # Update to new window
                    self.prev_window = current_window
                    self.prev_proc = current_proc
                    self.prev_details = current_details
                    self.prev_category = current_category
                    self.start_time = datetime.datetime.now()

                # Check idle status periodically
                idle_check_counter += 1
                if idle_check_counter >= idle_check_frequency:
                    idle_check_counter = 0
                    self._check_idle_status(current_window, current_proc, current_details, current_category)

            except Exception as e:
                print(f"Error in polling loop: {e}")
                time.sleep(window_check_interval)
            time.sleep(window_check_interval)

    def _check_idle_status(self, current_window, current_proc, current_details, current_category):
        """Check and handle idle status"""
        idle_seconds = self.get_idle_seconds()
        current_category = self.get_category(self.prev_window, self.prev_proc, self.prev_details)
        # 1 hour for meetings, 5 min for others
        idle_threshold = 3600 if current_category == "Meetings" else 300

        is_idle = idle_seconds >= idle_threshold

        if is_idle and not self.was_idle:
            # Just went idle
            print(f"Going idle after {idle_seconds:.1f} seconds")
            self.idle_start = datetime.datetime.now()

            # Log current activity before going idle
            if self.prev_window and self.start_time:
                duration = (self.idle_start - self.start_time).total_seconds()
                if duration >= 1:
                    self.log_activity(
                        self.start_time, self.idle_start,
                        self.prev_window, self.prev_proc,
                        self.prev_details, self.prev_category
                    )
                    print(f"Logged before idle: {self.prev_window} ({duration:.1f}s)")

            self.was_idle = True

        elif not is_idle and self.was_idle:
            # Just became active
            print("Becoming active from idle")

            if hasattr(self, 'idle_start') and self.idle_start:
                idle_duration = (datetime.datetime.now() - self.idle_start).total_seconds()
                if idle_duration >= 300:  # Log idle periods longer than 5 minutes
                    self.log_activity(
                        self.idle_start, datetime.datetime.now(),
                        "Inactive", "", "", "Inactive"
                    )
                    print(f"Logged idle period: {idle_duration:.1f} seconds")

            # Reset for new activity
            self.start_time = datetime.datetime.now()
            self.was_idle = False
            # Update current window info since we might have missed changes while idle
            self.prev_window = current_window
            self.prev_proc = current_proc
            self.prev_details = current_details
            self.prev_category = current_category

    def start(self):
        """Start the activity logger"""
        if not self.running:
            self.running = True
            self.thread = threading.Thread(target=self.logger_loop, daemon=True)
            self.thread.start()

    def stop(self):
        """Stop the activity logger"""
        if self.running:
            self.running = False
            self.uninstall_hook()
            if self.thread:
                self.thread.join(timeout=2)
            self.thread = None

    def restart(self):
        """Restart the activity logger"""
        self.stop()
        time.sleep(0.5)
        self.start()

    def open_log(self):
        """Open the log viewer"""
        from ui.viewer import LogViewer
        LogViewer(self.log_path)

    def update_historical_categories(self, key, new_category):
        """Update category for all historical entries of a key in ActivityLog.csv"""
        try:
            if not os.path.exists(self.log_path):  # <-- changed from self.log_file
                return 0
            
            # Read all rows
            with open(self.log_path, 'r', encoding='utf-8') as f:  # <-- changed from self.log_file
                reader = list(csv.reader(f))
            
            if not reader:
                return 0
                
            headers = reader[0]
            rows = reader[1:]

            # Find the ProcessName and Category columns, do not change to ApplicationKey
            try:
                process_name_index = headers.index('ProcessName')
                category_index = headers.index('Category')
            except ValueError:
                print("Could not find ProcessName or Category columns")
                return 0
            
            # Prepare key variations for matching
            key_variations = set()
            key_variations.add(key)  # Original key
            
            # If key has extension, add version without extension
            if '.' in key:
                key_without_ext = key.rsplit('.', 1)[0]
                key_variations.add(key_without_ext)
            else:
                # If key doesn't have extension, add common extensions
                common_extensions = ['.exe', '.com', '.bat', '.cmd', '.msi']
                for ext in common_extensions:
                    key_variations.add(key + ext)
            
            print(f"Looking for key variations: {key_variations}")
            
            # Update rows with matching ProcessName (using any variation)
            updated_count = 0
            matched_keys = set()
            
            for row in rows:
                if len(row) > max(process_name_index, category_index):
                    row_key = row[process_name_index]
                    
                    # Check if row key matches any of our key variations (case-insensitive)
                    key_match = False

                    # Direct match (case-insensitive)
                    if row_key.lower() in {k.lower() for k in key_variations}:
                        key_match = True
                        matched_keys.add(row_key)

                    # Also check if row_key without extension matches key without extension (case-insensitive)
                    if not key_match and '.' in row_key:
                        row_key_without_ext = row_key.rsplit('.', 1)[0].lower()
                        key_without_ext = key.rsplit('.', 1)[0].lower() if '.' in key else key.lower()
                        if row_key_without_ext == key_without_ext:
                            key_match = True
                            matched_keys.add(row_key)

                    # Only change the category, do not change the ProcessName
                    if key_match and row[category_index] != new_category:
                        row[category_index] = new_category
                        updated_count += 1
            
            print(f"Matched keys in CSV: {matched_keys}")
            
            # Write back to file only if changes were made
            if updated_count > 0:
                with open(self.log_path, 'w', newline='', encoding='utf-8') as f:  # <-- changed from self.log_file
                    writer = csv.writer(f)
                    writer.writerow(headers)
                    writer.writerows(rows)
                
                print(f"Updated {updated_count} historical entries for '{key}' to category '{new_category}'")
                
                # Update the app_categories for all matched keys
                for matched_key in matched_keys:
                    self.app_categories[matched_key] = new_category
                
                # Save updated categories
                self.save_app_categories()
                
                # Force regeneration of summary file
                self.generate_summary()
            
            return updated_count
            
        except Exception as e:
            print(f"Error updating historical categories: {e}")
            import traceback
            traceback.print_exc()
            return 0

    def notify_category_change(self, key, new_category):
        """Notify all open viewers about category change"""
        # Import here to avoid circular imports
        try:
            from ui.viewer import LogViewer
            
            # Update all open viewer instances
            for log_path, viewer in LogViewer._instances.items():
                try:
                    if viewer.root.winfo_exists():
                        # Schedule refresh on the main thread
                        viewer.root.after_idle(viewer.refresh_after_category_change)
                except:
                    pass
                    
        except ImportError:
            pass

    def get_process_key(self, process_name):
        """Get standardized process key (with or without extension based on configuration)"""
        # You can customize this logic based on how you want keys to be stored
        # For now, let's store with extension but allow matching without
        return process_name

    def normalize_key_for_matching(self, key1, key2):
        """Check if two keys match, ignoring extension and case"""
        # Remove extension and convert to lowercase
        key1_base = key1.rsplit('.', 1)[0].lower() if '.' in key1 else key1.lower()
        key2_base = key2.rsplit('.', 1)[0].lower() if '.' in key2 else key2.lower()
        return key1_base == key2_base

    def generate_summary(self):
        """Generate a summary CSV with total duration per process and category."""
        try:
            if not os.path.exists(self.log_path):
                print("No log file found for summary generation.")
                return

            summary = {}
            with open(self.log_path, 'r', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    process = row.get("ProcessName", "")
                    category = row.get("Category", "")
                    duration = int(row.get("DurationSeconds", 0))
                    key = (process, category)
                    summary[key] = summary.get(key, 0) + duration

            # Write summary to CSV
            summary_path = os.path.join(os.path.dirname(self.log_path), "ActivitySummary.csv")
            with open(summary_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(["ProcessName", "Category", "TotalDurationSeconds"])
                for (process, category), total_duration in summary.items():
                    writer.writerow([process, category, total_duration])

            print(f"Summary written to {summary_path}")

        except Exception as e:
            print(f"Error generating summary: {e}")