# ActivityLogger

ActivityLogger is a Windows desktop application that automatically tracks your computer activity by monitoring active windows and logging usage data to a CSV file. The application runs in the system tray and provides a simple interface for viewing, analyzing, and exporting your activity logs.

René Coffeng, 28 June 2025
Last edit: 22 July 2024
## Features

- **Automatic Activity Logging:** Monitors active windows and records application usage, window titles, and idle time.
- **System Tray Integration:** Runs quietly in the background with a tray icon for quick access to controls.
- **Session & Idle Tracking:** Tracks total logged time, session duration, and idle periods.
- **Log Viewer:** Built-in viewer to browse and summarize your activity data.
- **Exportable Logs:** Logs are saved as CSV files, compatible with Excel, Power BI, and other tools.
- **Custom Categories:** Supports categorization of activities for better analysis.
- **Startup Integration:** Optionally copies itself to the Windows Startup folder for automatic launch.
- **Help:** Built-in help viewer for usage tips.

## Usage

1. **Build the Application:**  
   Use the provided `build.py` script to build the executable with PyInstaller.
   ```sh
   python build.py
   ```
2. **Run ActivityLogger:**  
   Launch `ActivityLogger.exe`. The app will appear in the system tray.
3. **Access Controls:**  
   Right-click the tray icon to start/stop logging, open the log file, view help, or exit.
4. **View Logs:**  
   Use the built-in log viewer or open the CSV file in Excel for further analysis.

## Requirements

- Windows 10/11
- Python 3.8+ (for building)
- [PyInstaller](https://pyinstaller.org/)
- [Pillow](https://python-pillow.org/) (for icon generation)
- [psutil](https://pypi.org/project/psutil/) (for process checks)

## Building and Running

To build the executable, run:
```sh
python build.py
```
This will generate `ActivityLogger.exe` in the `dist` folder and optionally copy it to your Startup folder.

## License

This project is provided as-is for personal productivity and research purposes.

---

For more information, see the source code comments or use the built-in Help menu.

To run the application after building, simply execute ActivityLogger.exe from the dist directory. The application will start and place an icon in the system tray. Right-clicking the icon provides access to the application's functions, including starting/stopping logging, viewing the log, and exiting. + +## Project Structure + +The project is organized into a modular structure, separating core logic, UI, and system tray management. + +* main.py: The main entry point of the application. It initializes the ActivityLogger and the TrayManager, connecting the core logic to the user-facing tray interface. + +* build.py: A comprehensive build script that automates the creation of the final executable using PyInstaller. Its responsibilities include:

Automatically finding or generating an application icon (icon.ico).
Incrementing the build version number and writing it to version_info.txt.
Running PyInstaller with the necessary configurations, hidden imports, and data files.
Optionally signing the final .exe with signtool.
+* core/: This package contains the fundamental logic of the application.

logger.py: Likely contains the ActivityLogger class, which is responsible for monitoring the active window (win32gui), tracking application names and titles, calculating idle time, and writing the data to the CSV log file.
+* tray/: This package manages all system tray interactions.

tray_manager.py: Contains the TrayManager class. This class uses pystray to create the icon and the context menu (Start/Stop, View Log, etc.). It acts as the controller, translating user actions from the tray menu into calls to the ActivityLogger instance.
+* ui/: This package holds the graphical user interface components.

Based on the tkinter imports in the build script, this directory likely contains modules for the log viewer and help windows. A log_viewer.py module would be responsible for reading the CSV log and displaying it in a user-friendly, summarized format.
+* create_icon.py: (Inferred from build.py) A utility script that uses the Pillow library to programmatically generate the icon.ico file if one is not found. +


