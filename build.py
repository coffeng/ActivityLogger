"""
Build script for ActivityLogger.exe using PyInstaller
"""
import os
import sys
import subprocess
import shutil
import time
from pathlib import Path

def find_existing_icon():
    """Find existing icon files in ActivityLogger and related directories"""
    current_dir = Path(__file__).parent.absolute()
    
    # Look for icon files in various locations
    search_paths = [
        current_dir,  # ActivityLogger folder (current directory)
        current_dir / "assets",  # Assets subfolder
        current_dir / "icons",   # Icons subfolder
        current_dir.parent,      # Parent directory
        current_dir.parent / "ActivityLogger",  # Sibling ActivityLogger folder
    ]
    
    icon_names = ['icon.ico', 'ActivityLogger.ico', 'app.ico', 'stopwatch.ico']
    
    for search_path in search_paths:
        if not search_path.exists():
            continue
            
        for icon_name in icon_names:
            icon_path = search_path / icon_name
            if icon_path.exists():
                print(f"Found existing icon: {icon_path}")
                # Copy to ActivityLogger directory if not already there
                local_icon = current_dir / "icon.ico"
                if icon_path != local_icon:
                    try:
                        shutil.copy2(icon_path, local_icon)
                        print(f"Copied icon to: {local_icon}")
                        return True
                    except Exception as e:
                        print(f"Could not copy icon: {e}")
                        return False
                return True
    
    return False

def force_remove_directory(path):
    """Force remove directory with retry logic"""
    if not path.exists():
        return True
        
    max_attempts = 5
    for attempt in range(max_attempts):
        try:
            shutil.rmtree(path)
            print(f"Successfully removed {path}")
            return True
        except PermissionError as e:
            print(f"Attempt {attempt + 1}: Permission denied removing {path}")
            if attempt < max_attempts - 1:
                print("Retrying in 2 seconds...")
                time.sleep(2)
                
                # Try to change permissions
                try:
                    for root, dirs, files in os.walk(path):
                        for file in files:
                            file_path = os.path.join(root, file)
                            try:
                                os.chmod(file_path, 0o777)
                            except:
                                pass
                        for dir in dirs:
                            dir_path = os.path.join(root, dir)
                            try:
                                os.chmod(dir_path, 0o777)
                            except:
                                pass
                except:
                    pass
            else:
                print(f"Failed to remove {path} after {max_attempts} attempts")
                return False
        except Exception as e:
            print(f"Error removing {path}: {e}")
            return False
    
    return False

def check_for_running_processes():
    """Check if ActivityLogger.exe is running"""
    try:
        import psutil
        for proc in psutil.process_iter(['pid', 'name']):
            if proc.info['name'] and 'ActivityLogger.exe' in proc.info['name']:
                print(f"Warning: ActivityLogger.exe is running (PID: {proc.info['pid']})")
                return True
    except ImportError:
        pass
    except Exception:
        pass
    return False

def create_icon_if_missing():
    """Create high-quality icon in ActivityLogger folder if it doesn't exist"""
    # Icon should be in the ActivityLogger folder (current directory)
    current_dir = Path(__file__).parent.absolute()
    icon_path = current_dir / "icon.ico"
    
    # First try to find existing icon
    if find_existing_icon():
        return True
    
    if not icon_path.exists():
        print("No icon found, creating high-quality icon in ActivityLogger folder...")
        try:
            # Import and run the icon creation
            from create_icon import create_stopwatch_icon
            images = create_stopwatch_icon()
            if images:
                # Save icon in ActivityLogger folder
                images[0].save(str(icon_path), format='ICO', sizes=[(img.width, img.height) for img in images])
                print(f"High-quality icon created: {icon_path}")
                
                # Also save preview in ActivityLogger folder
                preview_path = current_dir / "icon_preview.png"
                images[-1].save(str(preview_path), format='PNG')
                print(f"Preview saved: {preview_path}")
                return True
            else:
                print("Failed to create icon")
                return False
        except ImportError:
            print("Could not import PIL. Install with: pip install pillow")
            return False
        except Exception as e:
            print(f"Error creating icon: {e}")
            return False
    return True

def copy_to_startup_and_run(exe_path):
    """Copy executable to Start Menu and optionally run it"""
    try:
        import shutil

        # Get the Startup folder path
        startup_folder = Path.home() / "AppData" / "Roaming" / "Microsoft" / "Windows" / "Start Menu" / "Programs" / "Startup"

        # Ensure startup folder exists
        startup_folder.mkdir(parents=True, exist_ok=True)

        # Define destination path
        startup_exe_path = startup_folder / "ActivityLogger.exe"

        # Copy the executable
        print(f"Copying to Startup folder: {startup_exe_path}")
        shutil.copy2(exe_path, startup_exe_path)
        print(f"Successfully copied ActivityLogger.exe to Startup folder")

        # Ask user if they want to run it
        print()
        response = input("Do you want to run ActivityLogger.exe now? (y/N): ").strip().lower()

        if response == 'y':
            print("Starting ActivityLogger.exe...")
            try:
                import subprocess
                # Launch via cmd to ensure it detaches and works from Startup
                subprocess.Popen(
                    ["cmd", "/c", "start", "", str(startup_exe_path)],
                    shell=True
                )
                print("ActivityLogger.exe started successfully!")
                time.sleep(2)  # Give it time to start
            except Exception as e:
                print(f"Error starting ActivityLogger.exe: {e}")
        else:
            print("ActivityLogger.exe not started. You can run it manually from the Startup folder.")

        return True

    except Exception as e:
        print(f"Error copying to Startup folder: {e}")
        return False

def main():
    """Build ActivityLogger.exe"""
    print("Building ActivityLogger.exe...")
    print(f"Working directory: {Path(__file__).parent.absolute()}")
    
    # Check for running processes first
    if check_for_running_processes():
        response = input("ActivityLogger is running. Continue anyway? (y/N): ")
        if response.lower() != 'y':
            return 1
    
    # Get the current directory (should be ActivityLogger folder)
    current_dir = Path(__file__).parent.absolute()
    
    # Define paths
    main_py = current_dir / "main.py"
    dist_dir = current_dir / "dist"
    build_dir = current_dir / "build"
    spec_file = current_dir / "ActivityLogger.spec"
    icon_path = current_dir / "icon.ico"
    
    # Check if main.py exists
    if not main_py.exists():
        print(f"Error: {main_py} not found!")
        return 1
    
    # Handle icon
    has_icon = create_icon_if_missing()
    
    # Skip cleaning directories for faster builds
    print("Skipping directory cleanup for faster builds...")
    
    # Skip removing dist folder - commented out
    # if dist_dir.exists():
    #     if not force_remove_directory(dist_dir):
    #         print("Could not remove dist directory. Continuing anyway...")
    
    # Skip removing build folder - commented out
    # if build_dir.exists():
    #     if not force_remove_directory(build_dir):
    #         print("Could not remove build directory. Continuing anyway...")
    
    if spec_file.exists():
        try:
            spec_file.unlink()
            print("Removed old spec file")
        except Exception as e:
            print(f"Could not remove spec file: {e}")
    
    # PyInstaller command
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",
        #"--clean",           # Remove temporary files before building
        "--noconfirm",       # Replace output directory without asking
        "--windowed",        # Hide console window (for GUI apps)
        "--name=ActivityLogger",
        "--add-data=core;core",
        "--add-data=ui;ui", 
        "--add-data=tray;tray",
        "--hidden-import=pystray._win32",
        "--hidden-import=PIL._tkinter_finder",
        "--hidden-import=tkinter",
        "--hidden-import=tkinter.ttk",
        "--hidden-import=tkinter.scrolledtext",
        "--hidden-import=win32gui",
        "--hidden-import=win32process",
        "--hidden-import=win32con",
        "--hidden-import=win32api",
        "--hidden-import=psutil",
        "--collect-all=pystray",
        "--collect-all=PIL",
    ]

    # Only use --icon if you trust your .ico file
    if has_icon and icon_path.exists():
        print("Using trusted icon file.")
        cmd.append(f"--icon={icon_path}")
    else:
        print("Building without icon (or icon not trusted/found).")

    cmd.append(str(main_py))
    
    print("Running PyInstaller...")
    
    try:
        result = subprocess.run(cmd, cwd=current_dir, check=True, capture_output=True, text=True)
        print("Build successful!")
        # ...existing code...

        exe_path = dist_dir / "ActivityLogger.exe"
        if exe_path.exists():
            print(f"ActivityLogger.exe created at: {exe_path}")
            print(f"File size: {exe_path.stat().st_size / 1024 / 1024:.2f} MB")

            # Sign the executable
            print("Signing ActivityLogger.exe...")
            signtool_cmd = [
                "signtool", "sign",
                "/tr", "http://timestamp.digicert.com",
                "/td", "SHA256",
                "/fd", "SHA256",
                "/sha1", "554cf2292aa90dcd2cda3326b39993c3407605ec",
                str(exe_path)
            ]
            try:
                sign_result = subprocess.run(signtool_cmd, capture_output=True, text=True)
                if sign_result.returncode == 0:
                    print("Successfully signed ActivityLogger.exe")
                else:
                    print("Failed to sign ActivityLogger.exe")
                    print("STDOUT:", sign_result.stdout)
                    print("STDERR:", sign_result.stderr)
            except Exception as e:
                print(f"Error running signtool: {e}")

            # Copy to Startup folder and ask to run
            print()
            print("=" * 60)
            copy_to_startup_and_run(exe_path)
            print("=" * 60)
           
        else:
            print("Warning: ActivityLogger.exe not found in dist directory")
            
    except subprocess.CalledProcessError as e:
        print(f"Build failed with return code {e.returncode}")
        print("STDOUT:", e.stdout)
        print("STDERR:", e.stderr)
        return 1
    
    except FileNotFoundError:
        print("Error: PyInstaller not found. Install with: pip install pyinstaller")
        return 1
    
    return 0

if __name__ == "__main__":
    sys.exit(main())