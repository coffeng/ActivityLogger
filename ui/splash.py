"""
Splash window utilities for Activity Logger
"""
import os
import sys
import datetime
import tkinter as tk
from tkinter import ttk
import configparser

from PIL import Image, ImageTk

from core.utils import ExeVersionInfo
from ui.help_viewer import HelpViewer


def _resource_path(rel_path: str) -> str:
    """
    Get absolute path to resource, works for dev and PyInstaller onefile.
    """
    base_dir = os.path.dirname(os.path.abspath(__file__))
    base_path = getattr(sys, '_MEIPASS', base_dir)
    # When frozen, ui/ is in _MEIPASS; icon may live next to executable
    # Try _MEIPASS first, then executable dir, then project dir
    candidates = [
        os.path.join(base_path, rel_path),
        os.path.join(
            os.path.dirname(sys.executable)
            if getattr(sys, 'frozen', False)
            else base_dir,
            rel_path,
        ),
        os.path.join(
            os.path.abspath(os.path.join(os.path.dirname(__file__), '..')),
            rel_path,
        ),
    ]
    for p in candidates:
        if os.path.exists(p):
            return p
    return rel_path  # Fallback; Tk will error if not found


def _center_window(win: tk.Toplevel | tk.Tk, width: int, height: int) -> None:
    sw = win.winfo_screenwidth()
    sh = win.winfo_screenheight()
    x = int((sw - width) / 2)
    y = int((sh - height) / 2)
    win.geometry(f"{width}x{height}+{x}+{y}")


def _build_splash_content(
    parent: tk.Toplevel | tk.Tk,
    log_path: str,
    include_help_button: bool = True,
):
    """Create inner widgets for the splash and return helpers."""
    container = ttk.Frame(parent, padding=12)
    container.pack(fill=tk.BOTH, expand=True)

    # Load icon image
    img_path = _resource_path('icon_preview.png')
    image = None
    try:
        pil_img = Image.open(img_path)
        # Scale to a max reasonable size for splash
        max_w, max_h = 180, 180
        pil_img.thumbnail((max_w, max_h), Image.LANCZOS)
        image = ImageTk.PhotoImage(pil_img)
    except Exception:
        image = None

    if image is not None:
        img_lbl = ttk.Label(container, image=image)
        img_lbl.image = image  # keep ref
        img_lbl.pack(pady=(4, 8))

    # Title / copyright
    title_lbl = ttk.Label(
        container,
        text="Activity Logger (c) 2025, by René Coffeng",
        font=('Arial', 12, 'bold'))
    title_lbl.pack(pady=(0, 4))

    # Build info from exe
    vi = ExeVersionInfo()
    version = vi.get_version() or 'dev'
    build_date = vi.get_build_date() or ''
    build_time = vi.get_build_time() or ''
    build_str = f"Build {version} {build_date} {build_time}".strip()
    build_lbl = ttk.Label(container, text=build_str, font=('Arial', 10))
    build_lbl.pack(pady=(0, 8))

    help_btn = None
    if include_help_button:
        help_btn = ttk.Button(container, text="Help", width=12)
        help_btn.pack(pady=(4, 0))

    return container, help_btn


def show_startup_splash(duration_ms: int = 5000):
    """
    Show the splash at application start. Blocks for duration using Tk.
    """
    # Determine default AppData log path (without importing heavier modules)
    try:
        user_home = os.path.expanduser('~')
        pc_name = os.getenv('COMPUTERNAME') or os.getenv('HOSTNAME') or ''
        if not pc_name:
            import socket as _socket
            pc_name = _socket.gethostname()
        base_dir = os.path.join(
            os.environ.get('LOCALAPPDATA', user_home), 'ActivityLogger'
        )
        os.makedirs(base_dir, exist_ok=True)
        log_path = os.path.join(base_dir, f"{pc_name}_ActivityLog.csv")
    except Exception:
        log_path = ''
    root = tk.Tk()
    root.withdraw()  # Hide base root; create a Toplevel as splash
    splash = tk.Toplevel(root)
    splash.title("Activity Logger")
    splash.attributes('-topmost', True)
    splash.resizable(False, False)

    # Build content on splash toplevel
    _, help_btn = _build_splash_content(
        splash, log_path, include_help_button=True
    )

    # Size to content, then center
    splash.update_idletasks()
    w = splash.winfo_reqwidth()
    h = splash.winfo_reqheight()
    _center_window(splash, w, h)

    # State
    state = {'help_open': False}

    def on_timeout():
        if not state['help_open']:
            try:
                splash.destroy()
            except Exception:
                pass
            try:
                root.quit()
                root.destroy()
            except Exception:
                pass

    def on_help_click():
        # Close splash immediately and mark help requested
        state['help_open'] = True
        try:
            splash.destroy()
        except Exception:
            pass
        try:
            root.quit()
            root.destroy()
        except Exception:
            pass

    if help_btn is not None:
        help_btn.configure(command=on_help_click)

    # Auto close after duration if help not opened
    root.after(duration_ms, on_timeout)
    try:
        root.mainloop()
    except Exception:
        # Fail-safe: ensure window is gone
        try:
            splash.destroy()
        except Exception:
            pass
        try:
            root.quit()
            root.destroy()
        except Exception:
            pass
    # Allow new Tk() later; return whether Help was requested and the path used
    try:
        import tkinter as _tk
        _tk._default_root = None
    except Exception:
        pass
    return state.get('help_open', False), log_path


def maybe_show_viewer_splash(
    parent: tk.Tk, log_path: str, duration_ms: int = 5000
):
    """
    Show splash for viewer once per day. Non-blocking.
    Stores date in %LOCALAPPDATA%\\ActivityLogger\\ActivityLogger.ini.
    """
    try:
        # Resolve INI path under AppData\Local\ActivityLogger
        user_home = os.path.expanduser('~')
        base_dir = os.path.join(
            os.environ.get('LOCALAPPDATA', user_home), 'ActivityLogger'
        )
        os.makedirs(base_dir, exist_ok=True)
        ini_path = os.path.join(base_dir, 'ActivityLogger.ini')

        cfg = configparser.ConfigParser()
        try:
            if os.path.exists(ini_path):
                cfg.read(ini_path, encoding='utf-8')
        except Exception:
            cfg = configparser.ConfigParser()

        if 'splash' not in cfg:
            cfg['splash'] = {}

        today = datetime.date.today().isoformat()
        last = cfg['splash'].get('viewer_last_date', '')
        if last == today:
            return  # Already shown today

    # Persist last shown date early to avoid duplicate splashes
        cfg['splash']['viewer_last_date'] = today
        try:
            with open(ini_path, 'w', encoding='utf-8') as f:
                cfg.write(f)
        except Exception:
            pass

        # Create as a transient top-level over the viewer parent
        splash = tk.Toplevel(parent)
        splash.title("Activity Logger")
        splash.attributes('-topmost', True)
        splash.resizable(False, False)
        splash.transient(parent)

        _, help_btn = _build_splash_content(
            splash, log_path, include_help_button=False
        )

        # Wire help button
        def on_help():
            try:
                HelpViewer(parent, log_path)
            except Exception:
                pass
            try:
                splash.destroy()
            except Exception:
                pass

        if help_btn is not None:
            try:
                help_btn.configure(command=on_help)
            except Exception:
                pass

        splash.update_idletasks()
        w = splash.winfo_reqwidth()
        h = splash.winfo_reqheight()
        _center_window(splash, w, h)

        def close_it():
            try:
                splash.destroy()
            except Exception:
                pass

        splash.after(duration_ms, close_it)
    except Exception:
        pass
