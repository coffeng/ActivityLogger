# ActivityLogger

ActivityLogger is a Windows desktop application that automatically tracks your computer activity by monitoring active windows and logging usage data to a CSV file. The application runs in the system tray and provides a simple interface for viewing, analyzing, and exporting your activity logs.

René Coffeng, 28 June 2025
Last edit: 01 Sep 2025
## Features

- **Automatic Activity Logging:** Monitors active windows and records application usage, window titles, and idle time.
- **System Tray Integration:** Runs quietly in the background with a tray icon for quick access to controls.
- **Session & Idle Tracking:** Tracks total logged time, session duration, and idle periods.
- **Log Viewer:** Built-in viewer to browse and summarize your activity data.
   - The viewer includes an Activity Log tab, a Summary tab, and a Category Graph tab.
   - Date range filters (Start / End) in the footer apply immediately when a date
      is selected. The viewer pauses automatic refresh while you interact with the
      date pickers so month navigation remains responsive.
   - Column sorting is supported on both the Activity and Summary lists by
      clicking the column header. Sorting is applied using identify-based header
      detection (works correctly when the window is resized) and persists across
      automatic refreshes so the view doesn't snap back to time order.
- **Exportable Logs:** Logs are saved as CSV files, compatible with Excel, Power BI, and other tools.
- **Custom Categories:** Supports categorization of activities for better analysis.
- **Startup Integration:** Optionally copies itself to the Windows Startup folder for automatic launch.
- **Help:** Built-in help viewer for usage tips.

## Usage

1. **Build the Application:**  
   Use the provided `build.py` script to build the executable with PyInstaller.
   ```powershell
   # Uses the system Python used for development (tested with Python 3.13)
   & 'C:\Path\To\Python313\python.exe' build.py
   ```
2. **Run ActivityLogger:**  
   Launch `ActivityLogger.exe`. The app will appear in the system tray.
3. **Access Controls:**  
   Right-click the tray icon to start/stop logging, open the log file, view help, or exit.
4. **View Logs:**  
   Use the built-in log viewer or open the CSV file in Excel for further analysis.

## Requirements

- Windows 10/11
 - Python 3.8+ (for building) — development and runtime in this repo were
    validated with Python 3.13 on Windows.
- [PyInstaller](https://pyinstaller.org/)
- [Pillow](https://python-pillow.org/) (for icon generation)
- [psutil](https://pypi.org/project/psutil/) (for process checks)

## Building and Running

To build the executable, run:
```sh
python build.py
```
This will generate `ActivityLogger.exe` in the `dist` folder and optionally copy it to your Startup folder.

## Building the installer (Inno Setup)

The repository includes an Inno Setup script (`ActivityLogger.iss`) and the
packaging workflow automates running the Inno compiler. The `setup.py`/build
pipeline will invoke Inno Setup (`ISCC.exe`) using the included `.iss` file so
you don't normally need to call the compiler yourself.

Prerequisites:
- Inno Setup installed on the build machine (provides `ISCC.exe`).

Create the installer via the repository's build script which orchestrates
PyInstaller and Inno Setup:
```powershell
# Uses the system Python used for development (tested with Python 3.13)
& 'C:\Path\To\Python313\python.exe' build.py
# or run the dedicated packaging step if present (some workflows use setup.py):
& 'C:\Path\To\Python313\python.exe' setup.py build_installer
```

The build pipeline will produce a `setup.exe` in the output folder configured
by the `.iss` script (commonly `setup/` or `output/`). The Inno script bundles
the `ActivityLogger.exe`, icons, and installer metadata defined in
`ActivityLogger.iss`.

Notes:
- You can edit `ActivityLogger.iss` to change installer metadata, default
   install path, or which files are included. Test generated installers in a VM
   before distributing.
- The build pipeline may optionally sign the installer using `signtool` if
   configured.

## Using setup.py (optional, development workflows)

If you'd like to work with the package using setuptools (editable install,
build a wheel, etc.), a `setup.py`-based workflow is supported. The examples
below use the full Python executable path (Windows PowerShell) so you can run
the exact interpreter used for development.

Install runtime and development dependencies (from `requirements.txt`):
```powershell
& 'C:\Path\To\Python313\python.exe' -m pip install -r requirements.txt
```

Install the project in editable/develop mode (useful during development):
```powershell
& 'C:\Path\To\Python313\python.exe' -m pip install -e .
# or, if a legacy setup.py workflow is required:
& 'C:\Path\To\Python313\python.exe' setup.py develop
```

Build a wheel or source distribution for release:
```powershell
& 'C:\Path\To\Python313\python.exe' -m pip wheel . -w dist
# or with setuptools directly:
& 'C:\Path\To\Python313\python.exe' setup.py sdist bdist_wheel
```

Install the built wheel from `dist/`:
```powershell
& 'C:\Path\To\Python313\python.exe' -m pip install dist\ActivityLogger-<version>-py3-none-any.whl
```

Notes:
- Editable installs (`pip install -e .`) let you edit source files in-place
   and instantly use changes without reinstalling.
- The repository's primary build path for end-users is the `build.py` script
   which wraps PyInstaller; `setup.py` is optional for Python packaging.

## License

This project is provided as-is for personal productivity and research purposes.

---

For more information, see the source code comments or use the built-in Help menu.


Automatically finding or generating an application icon (icon.ico).
Incrementing the build version number and writing it to version_info.txt.
Running PyInstaller with the necessary configurations, hidden imports, and data files.
Optionally signing the final .exe with signtool.
+* core/: This package contains the fundamental logic of the application.

logger.py: Likely contains the ActivityLogger class, which is responsible for monitoring the active window (win32gui), tracking application names and titles, calculating idle time, and writing the data to the CSV log file.
+* tray/: This package manages all system tray interactions.

tray_manager.py: Contains the TrayManager class. This class uses pystray to create the icon and the context menu (Start/Stop, View Log, etc.). It acts as the controller, translating user actions from the tray menu into calls to the ActivityLogger instance.
* ui/: This package holds the graphical user interface components. The log
   viewer (`ui/viewer.py`) contains the Activity Log tab, Summary tab, and
   Category Graph tab and implements the interactive behaviors described above.
+* create_icon.py: (Inferred from build.py) A utility script that uses the Pillow library to programmatically generate the icon.ico file if one is not found. +


