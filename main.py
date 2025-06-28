"""
Activity Logger - Main entry point
"""
import sys
import os

# Add the current directory to Python path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from core.logger import ActivityLogger
from tray.tray_manager import TrayManager


def main():
    logger = ActivityLogger()
    
    # Make logger accessible globally for the viewer
    import __main__
    __main__.logger_instance = logger
    
    # Create and start tray manager
    tray_manager = TrayManager(logger)
    tray_manager.run()


if __name__ == "__main__":
    main()