"""Views package for UI components."""

from .main_view import MainView
from .styles import COLORS, setup_styles
from .icons import create_app_icons, get_photo_images, setup_windows_taskbar

__all__ = [
    'MainView',
    'COLORS',
    'setup_styles',
    'create_app_icons',
    'get_photo_images',
    'setup_windows_taskbar'
]
