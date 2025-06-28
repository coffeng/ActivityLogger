# ActivityLogger

ActivityLogger is a Windows desktop application that automatically tracks your computer activity by monitoring active windows and logging usage data to a CSV file. The application runs in the system tray and provides a simple interface for viewing, analyzing, and exporting your activity logs.

René Coffeng, 28 June 2025

## Features

- **Automatic Activity Logging:** Monitors active windows and records application usage, window titles, and idle time.
- **System Tray Integration:** Runs quietly in the background with a tray icon for quick access to controls.
- **Session & Idle Tracking:** Tracks total logged time, session duration, and idle periods.
- **Log Viewer:** Built-in viewer to browse and summarize your activity data.
- **Exportable Logs:** Logs are saved as CSV files, compatible with Excel, Power BI, and other tools.
- **Custom Categories:** Supports categorization of activities for better analysis.
- **Startup Integration:** Optionally copies itself to the Windows Startup folder for automatic launch.
- **Help & Troubleshooting:** Built-in help viewer for usage tips and troubleshooting.

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

## Building

To build the executable, run:
```sh
python build.py
```
This will generate `ActivityLogger.exe` in the `dist` folder and optionally copy it to your Startup folder.

## License

This project is provided as-is for personal productivity and research purposes.

---

For more information, see the source code comments or use the built-in Help menu.
