"""
Activity Logger - Main entry point
"""
import sys
import os

# Add the current directory to Python path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from ui.splash import show_startup_splash


def main():
    # Show startup splash immediately (before opening/reading any log files)
    try:
        help_requested, help_log_path = show_startup_splash(duration_ms=5000)
    except Exception:
        # Even if splash fails, continue startup
        help_requested, help_log_path = (False, None)

    # If user clicked Help on splash, open help standalone now
    if help_requested:
        try:
            from ui.help_viewer import HelpViewer
            HelpViewer(help_log_path)
        except Exception:
            pass

    # Import heavy modules after splash so UI appears ASAP
    from core.logger import ActivityLogger
    from tray.tray_manager import TrayManager

    logger = ActivityLogger()

    # Make logger accessible globally for the viewer
    import __main__
    __main__.logger_instance = logger

    # Create and start tray manager
    tray_manager = TrayManager(logger)
    tray_manager.run()


if __name__ == "__main__":
    main()
