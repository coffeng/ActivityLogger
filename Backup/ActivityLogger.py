import sys
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
from pystray import Icon, Menu, MenuItem
from PIL import Image, ImageDraw
import math
import socket
import json
from collections import defaultdict

# Set log path - prefer OneDrive\Documents if available, otherwise use AppData\Local


def get_log_path():
    user_home = os.path.expanduser("~")
    pc_name = socket.gethostname()

    # Check for OneDrive folder
    onedrive_paths = [
        os.path.join(user_home, "OneDrive", "Documents"),
        os.path.join(user_home, "OneDrive - Personal", "Documents"),
        os.path.join(user_home, "OneDrive - GE HealthCare", "Documents"),
        # Add other common OneDrive patterns if needed
    ]

    for onedrive_path in onedrive_paths:
        if os.path.exists(onedrive_path):
            log_dir = os.path.join(onedrive_path, "ActivityLogger")
            if not os.path.exists(log_dir):
                os.makedirs(log_dir, exist_ok=True)
            return os.path.join(log_dir, f"{pc_name}_ActivityLog.csv")

    # Fallback to AppData\Local if OneDrive not found
    appdata_dir = os.path.join(os.environ.get(
        "LOCALAPPDATA", user_home), "ActivityLogger")
    if not os.path.exists(appdata_dir):
        os.makedirs(appdata_dir, exist_ok=True)
    return os.path.join(appdata_dir, f"{pc_name}_ActivityLog.csv")


LOG_PATH = get_log_path()
INTERVAL_SECONDS = 1

# Add these new imports for the hook
user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

# Define the hook procedure type
HOOKPROC = WINFUNCTYPE(ctypes.c_int, ctypes.c_int,
                       wintypes.WPARAM, wintypes.LPARAM)


class ActivityLogger:
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
        
        # Configuration file path - changed to CSV
        self.config_path = os.path.join(os.path.dirname(self.log_path), "ActivitySummary.csv")
        
        # Load app categories from config file
        self.app_categories = self.load_app_categories()
        
        # Track rows added since last config save
        self.rows_added_since_save = 0
        self.save_interval = 10  # Save config every 10 new rows

        # For idle detection, we still need polling but less frequent
        self.idle_check_interval = 5  # Check idle every 5 seconds instead of constantly

    def format_duration(self, total_seconds):
        """Format duration as dd hh:mm:ss"""
        days = int(total_seconds // 86400)
        hours = int((total_seconds % 86400) // 3600)
        minutes = int((total_seconds % 3600) // 60)
        seconds = int(total_seconds % 60)
        return f"{days:02d} {hours:02d}:{minutes:02d}:{seconds:02d}"

    def load_app_categories(self):
        """Load app categories from ActivitySummary.csv file with counts and durations"""
        default_categories = {
            "excel": "Work - Office",
            "winword": "Work - Office", 
            "powerpnt": "Work - Office",
            "outlook": "Email",
            "chrome": "Web Browsing",
            "firefox": "Web Browsing",
            "edge": "Web Browsing",
            "teams": "Meetings",
            "slack": "Communication",
            "zoom": "Meetings",
            "notepad": "Notes",
            "code": "Development",
            "pycharm": "Development",
            "cmd": "Terminal",
            "powershell": "Terminal"
        }
        
        if not os.path.exists(self.config_path):
            # Create initial CSV file
            self.save_app_categories(default_categories, {}, {})
            return default_categories
        
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                reader = csv.reader(f)
                
                # Skip comment lines and find headers
                headers = None
                for row in reader:
                    if row and not row[0].startswith('#') and row[0]:  # Non-empty, non-comment row
                        headers = row
                        break
                
                if not headers or 'Key' not in headers or 'Category' not in headers:
                    # Invalid or empty file, recreate with defaults
                    print("Invalid summary file format, using defaults")
                    self.save_app_categories(default_categories, {}, {})
                    return default_categories
                
                # Find column indices
                key_idx = headers.index('Key')
                category_idx = headers.index('Category')
                
                # Extract key-category mapping for internal use
                categories = {}
                for row in reader:
                    if len(row) > max(key_idx, category_idx):
                        key = row[key_idx].strip()
                        category = row[category_idx].strip()
                        if key and category:
                            categories[key] = category
            
            if categories:
                print(f"Loaded {len(categories)} categories from ActivitySummary.csv")
                # Merge with defaults for any missing keys
                merged_categories = default_categories.copy()
                merged_categories.update(categories)
                return merged_categories
            else:
                print("No valid categories found in summary file, using defaults")
                self.save_app_categories(default_categories, {}, {})
                return default_categories
                
        except Exception as e:
            print(f"Error loading ActivitySummary.csv: {e}")
            # Don't recreate file on error, just use defaults
            return default_categories

    def save_app_categories(self, categories=None, row_counts=None, durations=None):
        """Save app categories to ActivitySummary.csv file with statistics"""
        if categories is None:
            categories = self.app_categories
            
        # Get current statistics from log file
        if row_counts is None or durations is None:
            row_counts, durations = self.calculate_category_stats()
        
        # Find all keys that appear in the log file
        discovered_keys = set()
        if os.path.exists(self.log_path):
            try:
                with open(self.log_path, 'r', encoding='utf-8') as f:
                    reader = csv.reader(f)
                    headers = next(reader, None)
                    if headers:
                        try:
                            process_idx = headers.index('ProcessName')
                            window_idx = headers.index('WindowTitle') 
                            details_idx = headers.index('WindowDetails')
                            
                            for row in reader:
                                if len(row) > max(process_idx, window_idx, details_idx):
                                    process_name = row[process_idx].lower()
                                    window_title = row[window_idx].lower()
                                    window_details = row[details_idx].lower()
                                    
                                    # Extract potential keys from process names
                                    if process_name.endswith('.exe'):
                                        base_name = process_name[:-4]  # Remove .exe
                                        discovered_keys.add(base_name)
                                    
                                    # Extract potential keys from window titles
                                    words = window_title.split()
                                    for word in words:
                                        if len(word) > 3:  # Only consider words longer than 3 chars
                                            discovered_keys.add(word.lower())
                                    
                                    # Extract potential keys from window details
                                    words = window_details.split()
                                    for word in words:
                                        if len(word) > 3:
                                            discovered_keys.add(word.lower())
                                            
                        except (ValueError, IndexError):
                            pass
            except Exception as e:
                print(f"Error discovering keys: {e}")
        
        # Merge existing categories with discovered keys
        expanded_categories = categories.copy()
        
        # Add discovered keys that aren't already categorized
        for key in discovered_keys:
            if key not in expanded_categories:
                # Try to auto-categorize based on common patterns
                if any(x in key for x in ['excel', 'word', 'powerpoint', 'office']):
                    expanded_categories[key] = "Work - Office"
                elif any(x in key for x in ['chrome', 'firefox', 'edge', 'browser']):
                    expanded_categories[key] = "Web Browsing"
                elif any(x in key for x in ['teams', 'zoom', 'meet', 'skype']):
                    expanded_categories[key] = "Meetings"
                elif any(x in key for x in ['outlook', 'mail', 'thunderbird']):
                    expanded_categories[key] = "Email"
                elif any(x in key for x in ['code', 'studio', 'pycharm', 'eclipse']):
                    expanded_categories[key] = "Development"
                elif any(x in key for x in ['cmd', 'powershell', 'terminal', 'bash']):
                    expanded_categories[key] = "Terminal"
                elif any(x in key for x in ['notepad', 'note', 'text']):
                    expanded_categories[key] = "Notes"
                elif any(x in key for x in ['slack', 'discord', 'chat']):
                    expanded_categories[key] = "Communication"
                else:
                    expanded_categories[key] = "Uncategorized"
        
        # Create CSV data, but only include keys with non-zero counts
        csv_rows = []
        for key, category in expanded_categories.items():
            count = row_counts.get(key, 0)
            duration_seconds = durations.get(key, 0)
            
            # Only include keys that have been used (non-zero count)
            if count > 0:
                csv_rows.append({
                    'key': key,
                    'category': category,
                    'count': count,
                    'duration': self.format_duration(duration_seconds),
                    'duration_seconds': duration_seconds  # For sorting
                })
        
        # Sort by duration (highest first)
        csv_rows.sort(key=lambda x: x['duration_seconds'], reverse=True)
        
        # Write CSV file
        try:
            with open(self.config_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                
                # Write header with timestamp
                writer.writerow(['# ActivitySummary.csv'])
                writer.writerow(['# Last Updated:', datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
                writer.writerow(['# Total Entries:', len(csv_rows)])
                writer.writerow([])  # Empty row for separation
                
                # Write CSV headers
                writer.writerow(['Key', 'Category', 'Count', 'Duration'])
                
                # Write data rows
                for row in csv_rows:
                    writer.writerow([
                        row['key'],
                        row['category'], 
                        row['count'],
                        row['duration']
                    ])
                    
        except Exception as e:
            print(f"Error saving ActivitySummary.csv: {e}")

    def calculate_category_stats(self):
        """Calculate row counts and durations for each key from log file"""
        row_counts = defaultdict(int)
        durations = defaultdict(int)
        
        if not os.path.exists(self.log_path):
            return row_counts, durations
        
        try:
            with open(self.log_path, 'r', encoding='utf-8') as f:
                reader = csv.reader(f)
                headers = next(reader, None)
                if not headers:
                    return row_counts, durations
                
                # Find column indices
                try:
                    process_idx = headers.index('ProcessName')
                    window_idx = headers.index('WindowTitle') 
                    details_idx = headers.index('WindowDetails')
                    duration_idx = headers.index('DurationSeconds')
                except ValueError:
                    return row_counts, durations
                
                for row in reader:
                    if len(row) > max(process_idx, window_idx, details_idx, duration_idx):
                        process_name = row[process_idx].lower()
                        window_title = row[window_idx].lower()
                        window_details = row[details_idx].lower()
                        
                        try:
                            duration_seconds = int(row[duration_idx])
                        except (ValueError, IndexError):
                            continue
                        
                        # Check all possible keys (both predefined and discovered)
                        all_keys = set(self.app_categories.keys())
                        
                        # Add discovered keys from current row
                        if process_name.endswith('.exe'):
                            all_keys.add(process_name[:-4])
                        
                        # Check each key for matches
                        for key in all_keys:
                            if (key in process_name or 
                                key in window_title or 
                                key in window_details):
                                row_counts[key] += 1
                                durations[key] += duration_seconds
                                break  # Only count once per row
                                
        except Exception as e:
            print(f"Error calculating stats: {e}")
            
        return row_counts, durations

    def get_active_window_title(self):
        hwnd = win32gui.GetForegroundWindow()
        if hwnd == 0:
            return ""
        return win32gui.GetWindowText(hwnd)

    def get_active_process_name(self):
        hwnd = win32gui.GetForegroundWindow()
        if hwnd == 0:
            return ""
        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            return psutil.Process(pid).name()
        except Exception:
            return ""

    def get_window_details(self, window_title, process_name):
        # Extract the context/content part from the window title, not the app/process name
        try:
            # Remove process name or known suffixes from title
            proc = process_name.lower()
            title = window_title
            # Remove " - AppName" or "AppName - " patterns
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
                # Fallback: take the part before the last " - "
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
        class LASTINPUTINFO(ctypes.Structure):
            _fields_ = [("cbSize", ctypes.c_uint), ("dwTime", ctypes.c_uint)]
        lii = LASTINPUTINFO()
        lii.cbSize = ctypes.sizeof(LASTINPUTINFO)
        if ctypes.windll.user32.GetLastInputInfo(ctypes.byref(lii)):
            millis = win32api.GetTickCount() - lii.dwTime
            return millis / 1000.0
        return 0

    def log_activity(self, start, end, window, proc, details, category):
        # Duration in seconds, always positive
        duration = int((end - start).total_seconds())
        file_exists = os.path.isfile(self.log_path)
        with open(self.log_path, "a", newline='', encoding="utf-8") as f:
            writer = csv.writer(f)
            if not file_exists:
                writer.writerow([
                    "StartTime", "EndTime", "DurationSeconds", "WindowTitle", "WindowDetails", "ProcessName", "Category"
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

    def window_change_hook(self, nCode, wParam, lParam):
        """Hook procedure called when foreground window changes"""
        if nCode >= 0:
            # Process window change immediately
            self.process_window_change()

        # Call next hook
        return user32.CallNextHookEx(self.hook_id, nCode, wParam, lParam)

    def install_hook(self):
        """Install the window change hook - simplified approach"""
        try:
            # Let's try a simpler, more reliable approach using SetWinEventHook
            # Define the callback function properly
            def win_event_proc(hWinEventHook, event, hwnd, idObject, idChild, dwEventThread, dwmsEventTime):
                try:
                    if event == 0x0003:  # EVENT_SYSTEM_FOREGROUND
                        self.process_window_change()
                except Exception as e:
                    print(f"Hook callback error: {e}")

            # Create the callback type
            from ctypes import WINFUNCTYPE
            WinEventProc = WINFUNCTYPE(None, wintypes.HANDLE, wintypes.DWORD,
                                       wintypes.HWND, wintypes.LONG, wintypes.LONG,
                                       wintypes.DWORD, wintypes.DWORD)

            # Store the callback to prevent garbage collection
            self.hook = WinEventProc(win_event_proc)

            # Install the hook
            self.hook_id = user32.SetWinEventHook(
                0x0003,     # EVENT_SYSTEM_FOREGROUND
                0x0003,     # EVENT_SYSTEM_FOREGROUND
                None,       # hmodWinEventProc (NULL for out-of-context)
                self.hook,  # lpfnWinEventProc
                0,          # idProcess (0 for all processes)
                0,          # idThread (0 for all threads)
                0x0000      # dwFlags (WINEVENT_OUTOFCONTEXT)
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
                # Unhook the WinEvent hook
                result = user32.UnhookWinEvent(self.hook_id)
                if result:
                    print("WinEvent hook uninstalled successfully")
                else:
                    print("Failed to uninstall WinEvent hook")
                self.hook_id = None
            
            # Clear the hook callback reference
            if hasattr(self, 'hook'):
                self.hook = None
                
        except Exception as e:
            print(f"Error uninstalling hook: {e}")
  
    def process_window_change(self):
        """Process the window change event"""
        try:
            current_window = self.get_active_window_title()
            current_proc = self.get_active_process_name()
            current_details = self.get_window_details(
                current_window, current_proc)
            current_category = self.get_category(
                current_window, current_proc, current_details)

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
        """Simplified logger loop that should work reliably"""

        # Always use polling for now since hooks are problematic
        print("Using polling method for reliability")
        self.polling_fallback()
        return

        # The hook code below is disabled for now
        """
        # Try to install hook
        hook_installed = self.install_hook()
        
        if not hook_installed:
            print("Hook installation failed, using polling fallback")
            self.polling_fallback()
            return
        
        print("Hook installed successfully, using event-driven monitoring")
        
        # Start idle monitoring in separate thread
        idle_thread = threading.Thread(target=self.idle_monitor_loop, daemon=True)
        idle_thread.start()
        
        # Initialize current window
        self.prev_window = self.get_active_window_title()
        self.prev_proc = self.get_active_process_name()
        self.prev_details = self.get_window_details(self.prev_window, self.prev_proc)
        self.prev_category = self.get_category(self.prev_window, self.prev_proc, self.prev_details)
        self.start_time = datetime.datetime.now()
        
        # Simple message pump
        try:
            while self.running:
                time.sleep(0.1)  # Small sleep to prevent high CPU usage
        except Exception as e:
            print(f"Error in message loop: {e}")
        finally:
            self.uninstall_hook()
        """
    def uninstall_hook(self):
        """Uninstall the window change hook"""
        try:
            if hasattr(self, 'hook_id') and self.hook_id:
                # Unhook the WinEvent hook
                result = user32.UnhookWinEvent(self.hook_id)
                if result:
                    print("WinEvent hook uninstalled successfully")
                else:
                    print("Failed to uninstall WinEvent hook")
                self.hook_id = None
            
            # Clear the hook callback reference
            if hasattr(self, 'hook'):
                self.hook = None
                
        except Exception as e:
            print(f"Error uninstalling hook: {e}")

    def polling_fallback(self):
        """Enhanced polling method that definitely works"""
        print("Starting polling method")

        self.start_time = datetime.datetime.now()
        self.prev_window = self.get_active_window_title()
        self.prev_proc = self.get_active_process_name()
        self.prev_details = self.get_window_details(
            self.prev_window, self.prev_proc)
        self.prev_category = self.get_category(
            self.prev_window, self.prev_proc, self.prev_details)
        self.was_idle = False
        self.idle_start = None

        # More responsive polling
        window_check_interval = 0.5  # Check every 500ms for window changes
        idle_check_counter = 0
        idle_check_frequency = 10  # Check idle every 5 seconds (10 * 0.5s)

        print(f"Initial window: {self.prev_window}")

        while self.running:
            try:
                # Get current window info
                current_window = self.get_active_window_title()
                current_proc = self.get_active_process_name()
                current_details = self.get_window_details(
                    current_window, current_proc)
                current_category = self.get_category(
                    current_window, current_proc, current_details)

                # Check if window changed
                if (current_window != self.prev_window or
                    current_details != self.prev_details or
                    current_category != self.prev_category):

                    print(
                        f"Window changed: {self.prev_window} -> {current_window}")

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
                            print(
                                f"Logged activity: {self.prev_window} ({duration:.1f}s)")

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

                    idle_seconds = self.get_idle_seconds()
                    current_category = self.get_category(
                        self.prev_window, self.prev_proc, self.prev_details)
                    # 1 hour for meetings, 5 min for others
                    idle_threshold = 3600 if current_category == "Meetings" else 300

                    is_idle = idle_seconds >= idle_threshold

                    if is_idle and not self.was_idle:
                        # Just went idle
                        print(f"Going idle after {idle_seconds:.1f} seconds")
                        self.idle_start = datetime.datetime.now()

                        # Log current activity before going idle
                        if self.prev_window and self.start_time:
                            duration = (self.idle_start -
                                        self.start_time).total_seconds()
                            if duration >= 1:
                                self.log_activity(
                                    self.start_time, self.idle_start,
                                    self.prev_window, self.prev_proc,
                                    self.prev_details, self.prev_category
                                )
                                print(
                                    f"Logged before idle: {self.prev_window} ({duration:.1f}s)")

                        self.was_idle = True

                    elif not is_idle and self.was_idle:
                        # Just became active
                        print("Becoming active from idle")

                        if hasattr(self, 'idle_start') and self.idle_start:
                            idle_duration = (datetime.datetime.now(
                            ) - self.idle_start).total_seconds()
                            if idle_duration >= 300:  # Log idle periods longer than 5 minutes
                                self.log_activity(
                                    self.idle_start, datetime.datetime.now(),
                                    "Inactive", "", "", "Inactive"
                                )
                                print(
                                    f"Logged idle period: {idle_duration:.1f} seconds")

                        # Reset for new activity
                        self.start_time = datetime.datetime.now()
                        self.was_idle = False
                        # Update current window info since we might have missed changes while idle
                        self.prev_window = current_window
                        self.prev_proc = current_proc
                        self.prev_details = current_details
                        self.prev_category = current_category

            except Exception as e:
                print(f"Error in polling loop: {e}")
                time.sleep(window_check_interval)
            time.sleep(window_check_interval)

    def start(self):
        if not self.running:
            self.running = True
            self.thread = threading.Thread(
                target=self.logger_loop, daemon=True)
            self.thread.start()

    def stop(self):
        if self.running:
            self.running = False
            self.uninstall_hook()  # Make sure hook is removed
            if self.thread:
                self.thread.join(timeout=2)
            self.thread = None

    def restart(self):
        self.stop()
        time.sleep(0.5)
        self.start()

    def open_log(self):
        LogViewer(self.log_path)


def create_image():
    # Analogue stopwatch icon for tray
    img = Image.new('RGBA', (64, 64), (255, 255, 255, 0))
    d = ImageDraw.Draw(img)
    # Outer circle
    d.ellipse([8, 8, 56, 56], outline=(0, 0, 0),
              width=4, fill=(240, 240, 240, 255))
    # Stopwatch button
    d.rectangle([28, 2, 36, 14], fill=(0, 0, 0))
    # Ticks
    for angle in range(0, 360, 30):
        x1 = 32 + 22 * math.cos(math.radians(angle))
        y1 = 32 + 22 * math.sin(math.radians(angle))
        x2 = 32 + 26 * math.cos(math.radians(angle))
        y2 = 32 + 26 * math.sin(math.radians(angle))
        d.line([x1, y1, x2, y2], fill=(0, 0, 0), width=2)
    # Hands (fixed at 10:10)
    d.line([32, 32, 32, 16], fill=(200, 0, 0), width=4)  # minute
    d.line([32, 32, 44, 40], fill=(0, 0, 200), width=3)  # hour
    d.ellipse([30, 30, 34, 34], fill=(0, 0, 0))
    return img


class LogViewer:
    _instances = {}  # Class variable to track open instances

    def __init__(self, log_path):
        # Initialize is_duplicate first
        self.is_duplicate = False
        self.log_path = log_path
        
        # Set ActivitySummary.csv path
        self.summary_path = os.path.join(os.path.dirname(log_path), "ActivitySummary.csv")

        # Check if this log file is already being viewed
        if log_path in LogViewer._instances:
            existing_viewer = LogViewer._instances[log_path]
            try:
                if existing_viewer.root and existing_viewer.root.winfo_exists():
                    # Bring existing window to front
                    existing_viewer.root.lift()
                    existing_viewer.root.focus_force()
                    existing_viewer.root.attributes('-topmost', True)
                    existing_viewer.root.attributes('-topmost', False)
                    # Mark as duplicate and return early
                    self.is_duplicate = True
                    self.root = None
                    return
            except:
                # Window was destroyed, clean up
                del LogViewer._instances[log_path]

        # Register this instance
        LogViewer._instances[log_path] = self

        # Create the GUI
        import tkinter as tk
        from tkinter import ttk

        self.root = tk.Tk()
        self.root.title("Activity Log Viewer")
        self.root.geometry("900x600")  # Increased height for footer info
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        # Create notebook for tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # Create Activity Log tab
        self.activity_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.activity_frame, text="Activity Log")

        # Create Summary tab
        self.summary_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.summary_frame, text="Activity Summary")

        # Setup Activity Log tab
        self.setup_activity_tab()
        
        # Setup Summary tab
        self.setup_summary_tab()

        # Footer frame with buttons and stats (shared across tabs)
        footer_frame = tk.Frame(self.root, bg='lightgray', height=50)
        footer_frame.pack(side=tk.BOTTOM, fill=tk.X)
        footer_frame.pack_propagate(False)

        # Left side - Statistics labels (horizontally arranged)
        stats_frame = tk.Frame(footer_frame, bg='lightgray')

        # Statistics labels arranged horizontally (shortened labels)
        self.first_start_label = tk.Label(
            stats_frame, text="Started: --", bg='lightgray', font=('Arial', 8))
        self.first_start_label.pack(side=tk.LEFT, padx=(0, 15))

        self.last_stop_label = tk.Label(
            stats_frame, text="Last: --", bg='lightgray', font=('Arial', 8))
        self.last_stop_label.pack(side=tk.LEFT, padx=(0, 15))

        self.total_duration_label = tk.Label(
            stats_frame, text="Logged: --", bg='lightgray', font=('Arial', 8))
        self.total_duration_label.pack(side=tk.LEFT, padx=(0, 15))

        self.time_span_label = tk.Label(
            stats_frame, text="Session: --", bg='lightgray', font=('Arial', 8))
        self.time_span_label.pack(side=tk.LEFT, padx=(0, 15))

        # Add 5th statistic - Idle (Session - Logged) in seconds
        self.idle_time_label = tk.Label(
            stats_frame, text="Idle: --", bg='lightgray', font=('Arial', 8))
        self.idle_time_label.pack(side=tk.LEFT)

        # Center the stats frame vertically
        stats_frame.place(relx=0.02, rely=0.5, anchor=tk.W)

        # Right side - Buttons frame (vertically centered)
        buttons_frame = tk.Frame(footer_frame, bg='lightgray')

        # Recording status button
        self.recording_btn = tk.Button(
            buttons_frame, text="Start Logging", bg='blue', fg='white', command=self.toggle_recording)
        self.recording_btn.pack(side=tk.LEFT, padx=(0, 5))

        # Go To Folder button
        self.folder_btn = tk.Button(
            buttons_frame, text="Go To Folder", command=self.open_folder)
        self.folder_btn.pack(side=tk.LEFT)

        # Center the buttons frame vertically
        buttons_frame.place(relx=0.98, rely=0.5, anchor=tk.E)

        self.last_line_count = 0
        self.refresh_interval = 2000  # ms

        self.load_log()
        self.load_summary()
        self.update_recording_button()
        self.root.after(self.refresh_interval, self.refresh_data)

        # Start the main loop (only for new instances)
        self.root.mainloop()

    def setup_activity_tab(self):
        """Setup the Activity Log tab"""
        import tkinter as tk
        from tkinter import ttk

        # Main frame for treeview and scrollbars
        main_frame = tk.Frame(self.activity_frame)
        main_frame.pack(fill=tk.BOTH, expand=True)

        self.tree = ttk.Treeview(main_frame, show="headings")
        self.tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)

        self.scroll_y = ttk.Scrollbar(
            main_frame, orient="vertical", command=self.tree.yview)
        self.scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=self.scroll_y.set)

        self.scroll_x = ttk.Scrollbar(
            self.activity_frame, orient="horizontal", command=self.tree.xview)
        self.scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.configure(xscrollcommand=self.scroll_x.set)

    def setup_summary_tab(self):
        """Setup the Activity Summary tab"""
        import tkinter as tk
        from tkinter import ttk

        # Main frame for summary treeview and scrollbars
        summary_main_frame = tk.Frame(self.summary_frame)
        summary_main_frame.pack(fill=tk.BOTH, expand=True)

        # Summary info label
        self.summary_info_label = tk.Label(
            self.summary_frame, 
            text="Activity Summary - Categories sorted by total duration",
            font=('Arial', 10, 'bold'),
            pady=5
        )
        self.summary_info_label.pack(side=tk.TOP, fill=tk.X)

        self.summary_tree = ttk.Treeview(summary_main_frame, show="headings")
        self.summary_tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)

        # Bind right-click event to summary tree
        self.summary_tree.bind("<Button-3>", self.on_summary_right_click)

        self.summary_scroll_y = ttk.Scrollbar(
            summary_main_frame, orient="vertical", command=self.summary_tree.yview)
        self.summary_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.summary_tree.configure(yscrollcommand=self.summary_scroll_y.set)

        self.summary_scroll_x = ttk.Scrollbar(
            self.summary_frame, orient="horizontal", command=self.summary_tree.xview)
        self.summary_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.summary_tree.configure(xscrollcommand=self.summary_scroll_x.set)

    def on_summary_right_click(self, event):
        """Handle right-click on summary tree rows"""
        import tkinter as tk
        from tkinter import ttk
        
        # Get the item that was clicked
        item = self.summary_tree.identify_row(event.y)
        if not item:
            return
        
        # Get the values of the clicked row
        values = self.summary_tree.item(item, 'values')
        if not values or len(values) < 2:
            return
        
        key = values[0]  # First column is Key
        current_category = values[1]  # Second column is Category
        
        # Get all unique categories from the summary data
        categories = set()
        for child in self.summary_tree.get_children():
            child_values = self.summary_tree.item(child, 'values')
            if len(child_values) >= 2:
                categories.add(child_values[1])
        
        # Sort categories alphabetically
        sorted_categories = sorted(list(categories))
        
        # Create category selection window
        self.show_category_selector(event, key, current_category, sorted_categories)

    def show_category_selector(self, event, key, current_category, categories):
        """Show category selection popup with ability to add new categories"""
        import tkinter as tk
        from tkinter import ttk
        
        # Create popup window
        popup = tk.Toplevel(self.root)
        popup.title(f"Change Category for '{key}'")
        popup.geometry("350x500")  # Increased height for new category input
        popup.resizable(False, True)
        
        # Position popup near mouse cursor
        popup.geometry(f"+{event.x_root + 10}+{event.y_root + 10}")
        
        # Make popup modal
        popup.transient(self.root)
        popup.grab_set()
        
        # Header label
        header_label = tk.Label(
            popup, 
            text=f"Select category for: {key}",
            font=('Arial', 10, 'bold'),
            pady=10
        )
        header_label.pack()
        
        # Current category label
        current_label = tk.Label(
            popup,
            text=f"Current: {current_category}",
            font=('Arial', 9),
            fg='blue'
        )
        current_label.pack()
        
        # New category input frame
        new_category_frame = tk.Frame(popup)
        new_category_frame.pack(fill=tk.X, padx=10, pady=(5, 10))
        
        # New category label and entry
        new_category_label = tk.Label(
            new_category_frame,
            text="Add new category:",
            font=('Arial', 9)
        )
        new_category_label.pack(anchor=tk.W)
        
        new_category_entry = tk.Entry(
            new_category_frame,
            font=('Arial', 9),
            width=35
        )
        new_category_entry.pack(fill=tk.X, pady=(2, 0))
        
        # Create listbox with scrollbar
        list_label = tk.Label(
            popup,
            text="Or select existing category:",
            font=('Arial', 9)
        )
        list_label.pack(anchor=tk.W, padx=10)
        
        list_frame = tk.Frame(popup)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(5, 10))
        
        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        listbox = tk.Listbox(
            list_frame,
            yscrollcommand=scrollbar.set,
            font=('Arial', 9),
            selectmode=tk.SINGLE
        )
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=listbox.yview)
        
        # Populate listbox with categories
        for category in categories:
            listbox.insert(tk.END, category)
    
        # Find and select current category
        try:
            current_index = categories.index(current_category)
            listbox.selection_set(current_index)
            listbox.see(current_index)  # Scroll to make it visible
            listbox.activate(current_index)  # Set focus
        except ValueError:
            current_index = -1

        # Set focus based on whether current category was found
        if current_index >= 0:
            # Current category found in list - focus on listbox
            listbox.focus_set()
        else:
            # Current category not found in list - focus on entry field with current category pre-filled
            new_category_entry.insert(0, current_category)
            new_category_entry.select_range(0, tk.END)  # Select all text for easy editing
            new_category_entry.focus_set()
    
        # Function to handle selection from listbox
        def on_listbox_select(event):
            """When listbox item is selected, clear the new category entry"""
            if listbox.curselection():
                new_category_entry.delete(0, tk.END)
    
        # Function to handle typing in entry
        def on_entry_change(event):
            """When typing in entry, clear listbox selection"""
            listbox.selection_clear(0, tk.END)
    
        # Bind events
        listbox.bind('<<ListboxSelect>>', on_listbox_select)
        new_category_entry.bind('<KeyRelease>', on_entry_change)
    
        # Buttons frame
        button_frame = tk.Frame(popup)
        button_frame.pack(pady=10)
    
        def get_selected_category():
            """Get the selected category (either from listbox or entry)"""
            # Check if there's text in the new category entry
            new_category_text = new_category_entry.get().strip()
            if new_category_text:
                return new_category_text
        
            # Otherwise get from listbox selection
            selection = listbox.curselection()
            if selection:
                return categories[selection[0]]
        
            return None
    
        def on_change():
            selected_category = get_selected_category()
            if selected_category:
                self.change_category(key, selected_category)
                popup.destroy()
            else:
                # Show error message
                error_label = tk.Label(
                    popup,
                    text="Please select a category or enter a new one",
                    fg='red',
                    font=('Arial', 8)
                )
                error_label.pack()
                # Remove error after 3 seconds
                popup.after(3000, error_label.destroy)
    
        def on_cancel():
            popup.destroy()
    
        # Change button
        change_btn = tk.Button(
            button_frame,
            text="Change",
            command=on_change,
            bg='green',
            fg='white',
            width=12
        )
        change_btn.pack(side=tk.LEFT, padx=5)
    
        # Cancel button
        cancel_btn = tk.Button(
            button_frame,
            text="Cancel",
            command=on_cancel,
            width=12
        )
        cancel_btn.pack(side=tk.LEFT, padx=5)
    
        # Keyboard bindings
        def on_double_click(event):
            on_change()
    
        def on_key_press(event):
            if event.keysym == 'Return':
                on_change()
            elif event.keysym == 'Escape':
                on_cancel()
    
        listbox.bind('<Double-Button-1>', on_double_click)
        new_category_entry.bind('<Return>', lambda e: on_change())
        popup.bind('<Key>', on_key_press)
    
        # Set initial focus to new category entry for easy typing
        new_category_entry.focus_set()
    
        # If current category is selected in listbox, put it in the entry for editing
        if current_category in categories:
            new_category_entry.insert(0, current_category)
            new_category_entry.select_range(0, tk.END)  # Select all text for easy editing

    def change_category(self, key, new_category):
        """Change the category for a specific key in ActivitySummary.csv"""
        import __main__
        
        try:
            # Update the category in the logger's app_categories
            if hasattr(__main__, 'logger_instance'):
                logger = __main__.logger_instance
                logger.app_categories[key] = new_category
                
                # Force regeneration of ActivitySummary.csv with new category
                logger.save_app_categories()
                
                # Refresh the summary view to show changes
                self.load_summary()
                
                print(f"Changed category for '{key}' to '{new_category}'")
            
        except Exception as e:
            print(f"Error changing category: {e}")
            import tkinter.messagebox as messagebox
            messagebox.showerror("Error", f"Failed to change category: {e}")

    def load_summary(self):
        """Load ActivitySummary.csv data into the summary tab"""
        if not os.path.exists(self.summary_path):
            # Clear the summary tree if file doesn't exist
            self.summary_tree["columns"] = []
            self.summary_tree.delete(*self.summary_tree.get_children())
            self.summary_info_label.config(text="ActivitySummary.csv not found")
            return

        try:
            with open(self.summary_path, "r", encoding="utf-8") as f:
                reader = list(csv.reader(f))
                if not reader:
                    self.summary_info_label.config(text="ActivitySummary.csv is empty")
                    return

                # Skip comment lines and find headers
                data_start_idx = 0
                headers = []
                metadata = []
                
                for i, row in enumerate(reader):
                    if row and row[0].startswith('#'):
                        # Store metadata for display
                        if len(row) >= 2:
                            metadata.append(f"{row[0]} {row[1]}")
                        else:
                            metadata.append(row[0])
                        continue
                    elif row and not row[0].startswith('#') and row[0]:  # Non-empty, non-comment row
                        headers = row
                        data_start_idx = i + 1
                        break

                if not headers:
                    self.summary_info_label.config(text="No valid headers found in ActivitySummary.csv")
                    return

                # Get data rows
                rows = reader[data_start_idx:] if data_start_idx < len(reader) else []

                # Update info label with metadata
                info_text = "Activity Summary - Categories sorted by total duration"
                if metadata:
                    info_text += " | " + " | ".join(metadata)
                self.summary_info_label.config(text=info_text)

                # Set up columns
                self.summary_tree["columns"] = headers
                for col in headers:
                    self.summary_tree.heading(col, text=col)
                    # Set column widths based on content
                    if col == "Key":
                        self.summary_tree.column(col, width=120, anchor="w")
                    elif col == "Category":
                        self.summary_tree.column(col, width=150, anchor="w")
                    elif col == "Count":
                        self.summary_tree.column(col, width=80, anchor="center")
                    elif col == "Duration":
                        self.summary_tree.column(col, width=120, anchor="center")
                    else:
                        self.summary_tree.column(col, width=100, anchor="w")

                # Remove all old rows
                self.summary_tree.delete(*self.summary_tree.get_children())

                # Insert data rows (already sorted by duration in the file)
                for row in rows:
                    if row and any(cell.strip() for cell in row):  # Skip empty rows
                        # Pad row with empty strings if it's shorter than headers
                        padded_row = row + [''] * (len(headers) - len(row))
                        self.summary_tree.insert("", "end", values=padded_row[:len(headers)])

        except Exception as e:
            print(f"Error loading ActivitySummary.csv: {e}")
            self.summary_info_label.config(text=f"Error loading ActivitySummary.csv: {e}")

    def format_duration(self, total_seconds):
        """Format duration as dd hh:mm:ss"""
        days = int(total_seconds // 86400)
        hours = int((total_seconds % 86400) // 3600)
        minutes = int((total_seconds % 3600) // 60)
        seconds = int(total_seconds % 60)
        return f"{days:02d} {hours:02d}:{minutes:02d}:{seconds:02d}"

    def update_statistics(self, rows):
        """Update footer statistics based on log data since application start"""
        if not rows:
            self.first_start_label.config(text="Started: --")
            self.last_stop_label.config(text="Last: --")
            self.total_duration_label.config(text="Logged: --")
            self.time_span_label.config(text="Session: --")
            self.idle_time_label.config(text="Idle: --")
            return

        try:
            # Get the logger instance to find app start time
            import __main__
            if hasattr(__main__, 'logger_instance'):
                logger = __main__.logger_instance
                app_start_time = logger.app_start_time
            else:
                # Fallback to first entry if logger not available
                chronological_rows = rows[::-1]
                app_start_time = datetime.datetime.strptime(
                    chronological_rows[0][0], "%Y-%m-%d %H:%M:%S")

            # Filter rows to only include entries since app start
            session_rows = []

            for row in rows:
                try:
                    row_start_time = datetime.datetime.strptime(
                        row[0], "%Y-%m-%d %H:%M:%S")
                    if row_start_time >= app_start_time:
                        session_rows.append(row)
                except:
                    continue

            if not session_rows:
                # No entries since app started
                current_time = datetime.datetime.now()
                session_duration = (
                    current_time - app_start_time).total_seconds()

                self.first_start_label.config(
                    text=f"Started: {app_start_time.strftime('%Y-%m-%d %H:%M:%S')}")
                self.last_stop_label.config(text="Last: --")
                self.total_duration_label.config(text="Logged: 00 00:00:00")
                self.time_span_label.config(
                    text=f"Session: {self.format_duration(session_duration)}")
                self.idle_time_label.config(
                    text=f"Idle: {self.format_duration(session_duration)}")
                return

            # Get session statistics
            chronological_session_rows = session_rows[::-1]  # Reverse to chronological order

            # Last activity end time
            last_stop = datetime.datetime.strptime(
                chronological_session_rows[-1][1], "%Y-%m-%d %H:%M:%S")

            # Calculate total logged duration since app start (sum of DurationSeconds)
            total_logged_seconds = sum(
                int(row[2]) for row in session_rows if row[2].isdigit())

            # Calculate session time (time since app started)
            current_time = datetime.datetime.now()
            session_duration = (current_time - app_start_time).total_seconds()

            # Calculate idle time (session time - logged time)
            idle_time_seconds = session_duration - total_logged_seconds

            # Update labels - all times in dd hh:mm:ss format
            self.first_start_label.config(
                text=f"Started: {app_start_time.strftime('%Y-%m-%d %H:%M:%S')}")
            self.last_stop_label.config(
                text=f"Last: {last_stop.strftime('%Y-%m-%d %H:%M:%S')}")
            self.total_duration_label.config(
                text=f"Logged: {self.format_duration(total_logged_seconds)}")
            self.time_span_label.config(
                text=f"Session: {self.format_duration(session_duration)}")
            self.idle_time_label.config(
                text=f"Idle: {self.format_duration(idle_time_seconds)}")

        except Exception as e:
            print(f"Error updating statistics: {e}")

    def open_folder(self):
        """Open the folder containing the log file"""
        folder_path = os.path.dirname(self.log_path)
        os.startfile(folder_path)

    def toggle_recording(self):
        """Toggle logging on/off"""
        import __main__
        if hasattr(__main__, 'logger_instance'):
            logger = __main__.logger_instance
            if logger.running:
                logger.stop()
            else:
                logger.start()
        self.update_recording_button()

    def update_recording_button(self):
        """Update recording button text and color based on logging status"""
        import __main__
        if hasattr(__main__, 'logger_instance'):
            logger = __main__.logger_instance
            if logger.running:
                self.recording_btn.config(text="Logging", bg='red', fg='white')
            else:
                self.recording_btn.config(
                    text="Start Logging", bg='blue', fg='white')

    def on_close(self):
        """Handle window close event"""
        # Remove from instances when closed
        if not self.is_duplicate and self.log_path in LogViewer._instances:
            del LogViewer._instances[self.log_path]
        if self.root:
            self.root.destroy()

    def load_log(self):
        """Load activity log data"""
        if not os.path.exists(self.log_path):
            for col in self.tree["columns"]:
                self.tree.heading(col, text="")
            self.tree.delete(*self.tree.get_children())
            self.update_statistics([])
            return

        with open(self.log_path, "r", encoding="utf-8") as f:
            reader = list(csv.reader(f))
            if not reader:
                self.update_statistics([])
                return
            headers = reader[0]
            rows = reader[1:]

        # Reverse the rows so newest is on top
        rows = rows[::-1]

        # Filter out WindowDetails column from display (but keep in CSV)
        display_headers = [h for h in headers if h != "WindowDetails"]
        windowdetails_index = headers.index(
            "WindowDetails") if "WindowDetails" in headers else -1

        # Set up columns (without WindowDetails)
        self.tree["columns"] = display_headers
        for col in display_headers:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150, anchor="w")

        # Remove all old rows
        self.tree.delete(*self.tree.get_children())

        # Insert new rows (most recent first) - excluding WindowDetails column
        for row in rows:
            if windowdetails_index >= 0 and len(row) > windowdetails_index:
                # Remove WindowDetails value from display
                display_row = row[:windowdetails_index] + \
                    row[windowdetails_index + 1:]
            else:
                display_row = row
            self.tree.insert("", "end", values=display_row)

        self.last_line_count = len(rows)

        # Update statistics (using full rows with WindowDetails for calculations)
        self.update_statistics(rows)

    def refresh_data(self):
        """Refresh both Activity Log and Summary data"""
        try:
            # Refresh Activity Log
            if not os.path.exists(self.log_path):
                self.root.after(self.refresh_interval, self.refresh_data)
                return

            with open(self.log_path, "r", encoding="utf-8") as f:
                reader = list(csv.reader(f))
                rows = reader[1:] if len(reader) > 1 else []
            if len(rows) != self.last_line_count:
                self.load_log()

            # Refresh Summary (less frequently or when activity log changes)
            self.load_summary()

            # Update recording button status
            self.update_recording_button()
            
            # Schedule next refresh
            self.root.after(self.refresh_interval, self.refresh_data)
        except Exception as e:
            print(f"Error in refresh_data: {e}")
            # Still schedule next refresh even if there's an error
            self.root.after(self.refresh_interval, self.refresh_data)

class HelpViewer:
    def __init__(self, log_path):
        import tkinter as tk
        from tkinter import scrolledtext
        
        self.root = tk.Tk()
        self.root.title("Activity Logger Help")
        self.root.geometry("600x500")
        
        # Help text
        help_text = """
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
""" + log_path + """

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

def main():
    logger = ActivityLogger()
    # Make logger accessible globally for the viewer
    import __main__
    __main__.logger_instance = logger

    def on_start(icon, item):
        logger.start()

    def on_stop(icon, item):
        logger.stop()

    def on_restart(icon, item):
        logger.restart()

    def on_open_log(icon, item):
        logger.open_log()

    def on_help(icon, item):
        HelpViewer(logger.log_path)

    def on_exit(icon, item):
        logger.stop()
        # Close all LogViewer windows first
        for log_path, viewer in list(LogViewer._instances.items()):
            try:
                if viewer.root.winfo_exists():
                    viewer.root.destroy()
            except:
                pass
        LogViewer._instances.clear()
        icon.stop()

    # Put "Open Log File" at the top of the menu
    menu = Menu(
        MenuItem("Open Log File", on_open_log),
        MenuItem("Start Logging", on_start),
        MenuItem("Stop Logging", on_stop),
        MenuItem("Restart Logging", on_restart),
        MenuItem("Help", on_help),
        MenuItem("Exit", on_exit)
    )

    icon = Icon("Activity Logger", create_image(), "Activity Logger", menu)
    logger.start()
    icon.run()


if __name__ == "__main__":
    main()
