"""
System tray manager
"""
from pystray import Icon, Menu, MenuItem
from core.utils import create_tray_image
from ui.help_viewer import HelpViewer
from ui.viewer import LogViewer


class TrayManager:
    """Manages the system tray icon and menu"""
    
    def __init__(self, logger):
        self.logger = logger

    def run(self):
        """Start the tray icon"""
        def on_start(icon, item):
            self.logger.start()

        def on_stop(icon, item):
            self.logger.stop()

        def on_restart(icon, item):
            self.logger.restart()

        def on_open_log(icon, item):
            self.logger.open_log()

        def on_help(icon, item):
            HelpViewer(self.logger.log_path)

        def on_exit(icon, item):
            self.logger.stop()
            # Close all LogViewer windows first
            for log_path, viewer in list(LogViewer._instances.items()):
                try:
                    if viewer.root.winfo_exists():
                        viewer.root.destroy()
                except:
                    pass
            LogViewer._instances.clear()
            icon.stop()

        # Create menu with Open Log File at the top
        menu = Menu(
            MenuItem("Open Log File", on_open_log),
            MenuItem("Start Logging", on_start),
            MenuItem("Stop Logging", on_stop),
            MenuItem("Restart Logging", on_restart),
            MenuItem("Help", on_help),
            MenuItem("Exit", on_exit)
        )

        icon = Icon("Activity Logger", create_tray_image(), "Activity Logger", menu)
        self.logger.start()
        icon.run()