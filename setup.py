"""
Master build script for ActivityLogger.

This script performs the following steps:
1. Checks if the source code has been modified since the last .exe build.
2. If so, it runs build.py to create a new ActivityLogger.exe.
3. It then runs the Inno Setup compiler to create the final installer package.
"""
import os
import subprocess
import sys
from pathlib import Path

# --- Configuration ---
# Adjust this path if Inno Setup 6 is installed in a different location.
INNO_SETUP_COMPILER = Path("C:/Program Files (x86)/Inno Setup 6/ISCC.exe")

# --- Project Paths ---
ROOT_DIR = Path(__file__).parent.absolute()
EXE_PATH = ROOT_DIR / "dist" / "ActivityLogger.exe"
BUILD_SCRIPT = ROOT_DIR / "build.py"
ISS_FILE = ROOT_DIR / "ActivityLogger.iss"
SOURCE_DIRS = [
    ROOT_DIR,
    ROOT_DIR / "core",
    ROOT_DIR / "ui",
    ROOT_DIR / "tray",
]

def get_latest_source_mtime():
    """Find the most recent modification time of any relevant source file."""
    latest_mtime = 0
    # Check all .py files in the source directories
    for src_dir in SOURCE_DIRS:
        if not src_dir.exists():
            continue
        for py_file in src_dir.glob('**/*.py'):
            try:
                latest_mtime = max(latest_mtime, py_file.stat().st_mtime)
            except FileNotFoundError:
                continue  # File might have been deleted during glob

    # Also consider the build script itself as a source dependency
    if BUILD_SCRIPT.exists():
        latest_mtime = max(latest_mtime, BUILD_SCRIPT.stat().st_mtime)

    return latest_mtime

def is_rebuild_needed():
    """Checks if the executable is missing or outdated."""
    print("Checking if a rebuild is necessary...")
    if not EXE_PATH.exists():
        print(f"-> YES: Executable '{EXE_PATH}' does not exist.")
        return True

    exe_mtime = EXE_PATH.stat().st_mtime
    latest_source_mtime = get_latest_source_mtime()

    if latest_source_mtime > exe_mtime:
        print("-> YES: Source files have been modified since the last build.")
        return True

    print("-> NO: Executable is up-to-date.")
    return False

def run_build_script():
    """Executes the build.py script to create the .exe."""
    print("\n--- Running build script (build.py) ---")
    if not BUILD_SCRIPT.exists():
        print(f"Error: Build script '{BUILD_SCRIPT}' not found!")
        return False

    try:
        # Use sys.executable to ensure we use the same python interpreter
        subprocess.run([sys.executable, str(BUILD_SCRIPT)], check=True)
        print("--- Build script finished successfully. ---")
        return True
    except subprocess.CalledProcessError as e:
        print(f"Error: build.py failed with return code {e.returncode}.")
        return False
    except FileNotFoundError:
        print("Error: 'python' command not found. Make sure Python is in your PATH.")
        return False

def run_inno_setup():
    """Compiles the Inno Setup script."""
    print("\n--- Running Inno Setup Compiler ---")
    if not INNO_SETUP_COMPILER.exists():
        print(f"Error: Inno Setup Compiler not found at '{INNO_SETUP_COMPILER}'")
        print("Please install Inno Setup 6 or update the INNO_SETUP_COMPILER path in this script.")
        return False

    if not ISS_FILE.exists():
        print(f"Error: Inno Setup script '{ISS_FILE}' not found!")
        return False

    print(f"Compiling '{ISS_FILE}'...")
    try:
        subprocess.run([str(INNO_SETUP_COMPILER), str(ISS_FILE)], check=True)
        print("--- Inno Setup compilation finished successfully. ---")
        print(f"Installer created in '{ROOT_DIR / 'Output'}' directory.")
        return True
    except subprocess.CalledProcessError as e:
        print(f"Error: Inno Setup compilation failed with return code {e.returncode}.")
        return False

def main():
    """Main script execution."""
    if is_rebuild_needed():
        if not run_build_script():
            print("\nBuild failed. Aborting installer creation.")
            return 1  # Exit with error code

    if not run_inno_setup():
        print("\nInstaller creation failed.")
        return 1  # Exit with error code

    print("\nSetup process completed successfully!")
    return 0

if __name__ == "__main__":
    # Change to the script's directory to ensure relative paths work correctly
    os.chdir(Path(__file__).parent)
    sys.exit(main())