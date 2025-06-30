"""
Utility functions for Activity Logger
"""
import os
import socket
import math
from PIL import Image, ImageDraw


def get_log_path():
    """Get the appropriate log file path"""
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


def create_tray_image():
    """Create the system tray icon image"""
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


def format_duration(total_seconds):
    """Format duration as dd hh:mm:ss"""
    days = int(total_seconds // 86400)
    hours = int((total_seconds % 86400) // 3600)
    minutes = int((total_seconds % 3600) // 60)
    seconds = int(total_seconds % 60)
    return f"{days:02d} {hours:02d}:{minutes:02d}:{seconds:02d}"


class ExeVersionInfo:
    """Singleton-like class to read and cache version info from the EXE."""
    _instance = None

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance._initialized = False
        return cls._instance

    def __init__(self):
        if self._initialized:
            return
        import pefile
        import os
        import sys

        # Use the running executable if frozen, else use dist/ActivityLogger.exe
        if getattr(sys, 'frozen', False):
            exe_path = sys.executable
        else:
            exe_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'dist', 'ActivityLogger.exe'))

        self.version = "dev"
        self.build_date = ""
        self.build_time = ""
        if os.path.exists(exe_path):
            try:
                pe = pefile.PE(exe_path)
                if hasattr(pe, 'FileInfo'):
                    for fileinfo in pe.FileInfo:
                        if isinstance(fileinfo, list):
                            for subinfo in fileinfo:
                                if hasattr(subinfo, 'Key') and subinfo.Key == b'StringFileInfo':
                                    for st in subinfo.StringTable:
                                        entries = st.entries
                                        entries = {k.decode() if isinstance(k, bytes) else k:
                                                   v.decode() if isinstance(v, bytes) else v
                                                   for k, v in entries.items()}
                                        self.version = entries.get('FileVersion', 'dev')
                                        self.build_date = entries.get('BuildDate', '')
                                        self.build_time = entries.get('BuildTime', '')
                        elif hasattr(fileinfo, 'Key') and fileinfo.Key == b'StringFileInfo':
                            for st in fileinfo.StringTable:
                                entries = st.entries
                                entries = {k.decode() if isinstance(k, bytes) else k:
                                           v.decode() if isinstance(v, bytes) else v
                                           for k, v in entries.items()}
                                self.version = entries.get('FileVersion', 'dev')
                                self.build_date = entries.get('BuildDate', '')
                                self.build_time = entries.get('BuildTime', '')
            except Exception:
                pass
        self._initialized = True

    def get_version(self):
        return self.version

    def get_build_date(self):
        return self.build_date

    def get_build_time(self):
        return self.build_time