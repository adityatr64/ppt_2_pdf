"""
UI styling configuration.
Defines colors and ttk styles for the application.

⚠️ YEP YOU GUESSED IT ⚠️
⚠️ THIS PART IS ENTIRELY MADE WITH AI ⚠️
⚠️ I WOULD RATHER HANG MYSELF THAN WRITE FRONTEND IN PYTHON ⚠️

"""

from tkinter import ttk
from typing import Dict, Optional


# Application color palette
COLORS: Dict[str, str] = {
    'bg': '#f5f5f5',
    'card': '#ffffff',
    'primary': '#2962ff',
    'primary_hover': '#1e4bd8',
    'danger': '#dc3545',
    'success': '#28a745',
    'text': '#333333',
    'text_light': '#666666',
    'border': '#e0e0e0',
    'disabled': '#999999'
}


def setup_styles(colors: Optional[Dict[str, str]] = None) -> ttk.Style:
    """
    Configure ttk styles for a modern look.
    
    Args:
        colors: Optional color dictionary, defaults to COLORS
    
    Returns:
        Configured ttk.Style object
    """
    if colors is None:
        colors = COLORS
    
    style = ttk.Style()
    style.theme_use('clam')
    
    # Frame styles
    style.configure('Card.TFrame', background=colors['card'])
    style.configure('TFrame', background=colors['bg'])
    
    # Label styles
    style.configure(
        'Title.TLabel', 
        font=('Segoe UI', 16, 'bold'),
        background=colors['bg'],
        foreground=colors['primary']
    )
    
    style.configure(
        'TLabel',
        font=('Segoe UI', 10),
        background=colors['bg'],
        foreground=colors['text']
    )
    
    style.configure(
        'Status.TLabel',
        font=('Segoe UI', 9),
        background=colors['bg'],
        foreground=colors['text_light']
    )
    
    # Button styles
    style.configure(
        'TButton',
        font=('Segoe UI', 10),
        padding=(12, 6)
    )
    
    style.configure(
        'Primary.TButton',
        font=('Segoe UI', 11, 'bold'),
        padding=(20, 10)
    )
    
    style.map(
        'Primary.TButton',
        background=[
            ('active', colors['primary_hover']),
            ('!active', colors['primary'])
        ],
        foreground=[
            ('active', 'white'),
            ('!active', 'white')
        ]
    )
    
    # LabelFrame style
    style.configure(
        'Card.TLabelframe',
        background=colors['card'],
        borderwidth=1,
        relief='solid'
    )
    style.configure(
        'Card.TLabelframe.Label',
        font=('Segoe UI', 10, 'bold'),
        background=colors['card'],
        foreground=colors['text']
    )
    
    # Checkbutton style
    style.configure(
        'TCheckbutton',
        font=('Segoe UI', 10),
        background=colors['card']
    )
    
    # Progressbar style
    style.configure(
        'Custom.Horizontal.TProgressbar',
        troughcolor=colors['border'],
        background=colors['primary'],
        thickness=8
    )
    
    return style
