"""Application icon utilities.

⚠️ THIS PART IS ENTIRELY MADE WITH AI ⚠️
"""

import sys
import tkinter as tk

try:
    from PIL import Image, ImageTk
    HAS_PIL = True
except ImportError:
    HAS_PIL = False


def setup_windows_taskbar():
    """Set Windows taskbar icon identity (must be called before tkinter window creation)."""
    if sys.platform == 'win32':
        import ctypes
        # Set app user model ID so Windows shows our icon, not Python's
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID('ppt2pdf.converter.app.1')


def get_app_icon_path() -> str:
    """Return absolute path to assets/image.ico for dev and bundled runs."""
    base_path = getattr(sys, "_MEIPASS", os.path.abspath("."))
    return os.path.join(base_path, "assets", "image.ico")


def apply_window_icon(window: tk.Tk):
    """Apply assets/image.ico to a Tk/CTk window and return icon image refs."""
    icon_path = get_app_icon_path()
    if not os.path.exists(icon_path):
        return []

    try:
        window.iconbitmap(default=icon_path)
    except Exception:
        try:
            window.iconbitmap(icon_path)
        except Exception:
            pass

    if not HAS_PIL:
        return []

    try:
        icon_image = Image.open(icon_path) # pyright: ignore[reportPossiblyUnboundVariable]
        photo = ImageTk.PhotoImage(icon_image) # pyright: ignore[reportPossiblyUnboundVariable]
        window.iconphoto(True, photo) # pyright: ignore[reportArgumentType]
        return [photo]
    except Exception:
        return []
