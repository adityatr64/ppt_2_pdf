"""Views package for UI components."""

from .main_view import MainView
from .styles import COLORS, setup_styles
from .icons import apply_window_icon, get_app_icon_path, setup_windows_taskbar

__all__ = [
    'MainView',
    'COLORS',
    'setup_styles',
    'apply_window_icon',
    'get_app_icon_path',
    'setup_windows_taskbar'
]
