"""
Icon generation utilities.
Creates application icons programmatically using PIL.

⚠️ THIS PART IS ENTIRELY MADE WITH AI ⚠️

"""





import sys

# Try to import PIL for icon creation
try:
    from PIL import Image, ImageDraw, ImageTk
    HAS_PIL = True
except ImportError:
    HAS_PIL = False


def setup_windows_taskbar():
    """Set Windows taskbar icon identity (must be called before tkinter window creation)."""
    if sys.platform == 'win32':
        import ctypes
        # Set app user model ID so Windows shows our icon, not Python's
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID('ppt2pdf.converter.app.1')


def create_app_icons():
    """
    Create PDF icon in multiple sizes for taskbar and title bar.
    
    Returns:
        List of PIL Image objects in various sizes, or None if PIL is not available
    """
    if not HAS_PIL:
        return None
    
    icons = []
    for size in [16, 32, 48, 64, 128, 256]:
        img = Image.new('RGBA', (size, size), (0, 0, 0, 0)) # pyright: ignore[reportPossiblyUnboundVariable]
        draw = ImageDraw.Draw(img) # pyright: ignore[reportPossiblyUnboundVariable]
        
        # Scale factor
        s = size / 64
        
        # Document shape (white with blue border)
        doc_color = (255, 255, 255, 255)
        border_color = (41, 98, 255, 255)  # Blue
        accent_color = (220, 53, 69, 255)  # Red for PDF
        
        # Main document rectangle
        draw.rectangle(
            [int(8*s), int(4*s), int(52*s), int(60*s)], 
            fill=doc_color, 
            outline=border_color, 
            width=max(1, int(2*s))
        )
        
        # Folded corner
        draw.polygon(
            [(int(38*s), int(4*s)), (int(52*s), int(18*s)), (int(38*s), int(18*s))], 
            fill=(200, 200, 200, 255), 
            outline=border_color
        )
        
        # PDF badge
        draw.rectangle([int(14*s), int(36*s), int(46*s), int(52*s)], fill=accent_color)
        
        # Lines representing text
        draw.rectangle([int(14*s), int(22*s), int(34*s), int(24*s)], fill=(180, 180, 180, 255))
        draw.rectangle([int(14*s), int(28*s), int(40*s), int(30*s)], fill=(180, 180, 180, 255))
        
        icons.append(img)
    
    return icons


def get_photo_images(icons):
    """
    Convert PIL Images to Tkinter PhotoImages.
    
    Args:
        icons: List of PIL Image objects
    
    Returns:
        List of ImageTk.PhotoImage objects, or empty list if icons is None
    """
    if not icons or not HAS_PIL:
        return []
    
    return [ImageTk.PhotoImage(icon) for icon in icons] # pyright: ignore[reportPossiblyUnboundVariable]
