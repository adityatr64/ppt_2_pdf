"""
Main application view.
Handles all UI components and layout.

⚠️ THIS PART IS ENTIRELY MADE WITH AI ⚠️
⚠️ DON'T BLAME ME IF THE VIEW DIES I FUCKING HATE FRONTEND ⚠️

"""

import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Callable, List, Optional

from .styles import COLORS, setup_styles
from .icons import create_app_icons, get_photo_images


class MainView:
    """Main application window view."""
    
    def __init__(self, root: tk.Tk):
        self.root = root
        self.colors = COLORS
        self.icon_photos = []  # Keep reference to prevent garbage collection
        
        # Callbacks set by controller
        self.on_add_files: Optional[Callable[[], None]] = None
        self.on_remove_selected: Optional[Callable[[], None]] = None
        self.on_clear_all: Optional[Callable[[], None]] = None
        self.on_move_up: Optional[Callable[[], None]] = None
        self.on_move_down: Optional[Callable[[], None]] = None
        self.on_convert: Optional[Callable[[], None]] = None
        self.on_convert_separate: Optional[Callable[[], None]] = None
        self.on_drag_reorder: Optional[Callable[[int, int], None]] = None
        
        # Drag state
        self.drag_data = {'index': None}
        
        self._setup_window()
        self._setup_styles()
        self._setup_ui()
    
    def _setup_window(self):
        """Configure the main window."""
        self.root.title("PPT to PDF Converter & Merger")
        self.root.geometry("750x550")
        self.root.minsize(650, 450)
        self.root.configure(bg=self.colors['bg'])
        
        # Set app icon
        icons = create_app_icons()
        if icons:
            self.icon_photos = get_photo_images(icons)
            if self.icon_photos:
                self.root.iconphoto(True, *self.icon_photos) # pyright: ignore[reportArgumentType]
    
    def _setup_styles(self):
        """Configure ttk styles."""
        setup_styles(self.colors)
    
    def _setup_ui(self):
        """Create all UI components."""
        # Main frame with padding
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title label
        title_label = ttk.Label(
            main_frame, 
            text="PowerPoint to PDF Converter & Merger", 
            style='Title.TLabel'
        )
        title_label.pack(pady=(0, 15))
        
        # File list section
        self._setup_file_list(main_frame)
        
        # Buttons section
        self._setup_buttons(main_frame)
        
        # Options section
        self._setup_options(main_frame)
        
        # Progress section
        self._setup_progress(main_frame)
        
        # Convert button
        self._setup_convert_button(main_frame)
    
    def _setup_file_list(self, parent):
        """Create the file list panel."""
        list_frame = ttk.LabelFrame(
            parent, 
            text="  Selected Files  ", 
            style='Card.TLabelframe', 
            padding="10"
        )
        list_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Hint label
        hint_label = ttk.Label(
            list_frame, 
            text="Drag items to reorder or use the buttons below",
            style='Status.TLabel'
        )
        hint_label.configure(background=self.colors['card'])
        hint_label.pack(anchor=tk.W, pady=(0, 5))
        
        # Listbox with scrollbar
        list_container = ttk.Frame(list_frame, style='Card.TFrame')
        list_container.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(list_container)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.file_listbox = tk.Listbox(
            list_container, 
            selectmode=tk.SINGLE, 
            yscrollcommand=scrollbar.set,
            font=('Segoe UI', 10),
            activestyle='none',
            selectbackground=self.colors['primary'],
            selectforeground='white',
            borderwidth=1,
            relief='solid',
            highlightthickness=0
        )
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.file_listbox.yview)
        
        # Bind drag and drop
        self.file_listbox.bind('<Button-1>', self._on_click)
        self.file_listbox.bind('<B1-Motion>', self._on_drag)
        self.file_listbox.bind('<ButtonRelease-1>', self._on_drop)
    
    def _setup_buttons(self, parent):
        """Create file management buttons."""
        file_buttons_frame = ttk.Frame(parent)
        file_buttons_frame.pack(fill=tk.X, pady=10)
        
        # Left side buttons
        left_btns = ttk.Frame(file_buttons_frame)
        left_btns.pack(side=tk.LEFT)
        
        ttk.Button(
            left_btns, 
            text="+ Add Files", 
            command=lambda: self._invoke(self.on_add_files)
        ).pack(side=tk.LEFT, padx=2)
        
        ttk.Button(
            left_btns, 
            text="Remove", 
            command=lambda: self._invoke(self.on_remove_selected)
        ).pack(side=tk.LEFT, padx=2)
        
        ttk.Button(
            left_btns, 
            text="Clear All", 
            command=lambda: self._invoke(self.on_clear_all)
        ).pack(side=tk.LEFT, padx=2)
        
        # Right side buttons (reorder)
        right_btns = ttk.Frame(file_buttons_frame)
        right_btns.pack(side=tk.RIGHT)
        
        ttk.Button(
            right_btns, 
            text="Move Up", 
            command=lambda: self._invoke(self.on_move_up)
        ).pack(side=tk.LEFT, padx=2)
        
        ttk.Button(
            right_btns, 
            text="Move Down", 
            command=lambda: self._invoke(self.on_move_down)
        ).pack(side=tk.LEFT, padx=2)
    
    def _setup_options(self, parent):
        """Create options panel."""
        options_frame = ttk.LabelFrame(
            parent, 
            text="  Options  ", 
            style='Card.TLabelframe', 
            padding="10"
        )
        options_frame.pack(fill=tk.X, pady=5)
        
        self.delete_temp_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            options_frame, 
            text="Delete temporary PDF files after merging", 
            variable=self.delete_temp_var
        ).pack(anchor=tk.W, pady=2)
        
        self.open_after_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            options_frame, 
            text="Open PDF after conversion", 
            variable=self.open_after_var
        ).pack(anchor=tk.W, pady=2)
    
    def _setup_progress(self, parent):
        """Create progress bar and status label."""
        progress_frame = ttk.Frame(parent)
        progress_frame.pack(fill=tk.X, pady=10)
        
        self.status_label = ttk.Label(
            progress_frame, 
            text="Ready - Add files to get started", 
            style='Status.TLabel'
        )
        self.status_label.pack(anchor=tk.W, pady=(0, 5))
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            progress_frame, 
            variable=self.progress_var, 
            maximum=100, 
            mode='determinate',
            style='Custom.Horizontal.TProgressbar'
        )
        self.progress_bar.pack(fill=tk.X)
    
    def _setup_convert_button(self, parent):
        """Create the main convert buttons."""
        btn_frame = ttk.Frame(parent)
        btn_frame.pack(pady=15)
        
        # Merge button
        self.convert_btn = tk.Button(
            btn_frame, 
            text="Convert & Merge to PDF", 
            command=lambda: self._invoke(self.on_convert),
            font=('Segoe UI', 11, 'bold'),
            bg=self.colors['primary'],
            fg='white',
            activebackground=self.colors['primary_hover'],
            activeforeground='white',
            relief='flat',
            padx=30,
            pady=10,
            cursor='hand2'
        )
        self.convert_btn.pack(side=tk.LEFT, padx=5)
        
        # Separate PDFs button
        self.convert_separate_btn = tk.Button(
            btn_frame, 
            text="Make Separate PDFs", 
            command=lambda: self._invoke(self.on_convert_separate),
            font=('Segoe UI', 11, 'bold'),
            bg=self.colors['success'],
            fg='white',
            activebackground='#1e7b34',
            activeforeground='white',
            relief='flat',
            padx=30,
            pady=10,
            cursor='hand2'
        )
        self.convert_separate_btn.pack(side=tk.LEFT, padx=5)
        
        # Hover effects for merge button
        self.convert_btn.bind(
            '<Enter>', 
            lambda e: self.convert_btn.config(bg=self.colors['primary_hover'])
        )
        self.convert_btn.bind(
            '<Leave>', 
            lambda e: self.convert_btn.config(bg=self.colors['primary'])
        )
        
        # Hover effects for separate button
        self.convert_separate_btn.bind(
            '<Enter>', 
            lambda e: self.convert_separate_btn.config(bg='#1e7b34')
        )
        self.convert_separate_btn.bind(
            '<Leave>', 
            lambda e: self.convert_separate_btn.config(bg=self.colors['success'])
        )
    
    def _invoke(self, callback):
        """Safely invoke a callback if set."""
        if callback:
            callback()
    
    def _on_click(self, event):
        """Record the item being clicked for drag operation."""
        self.drag_data['index'] = self.file_listbox.nearest(event.y)
    
    def _on_drag(self, event):
        """Handle drag motion."""
        current_index = self.file_listbox.nearest(event.y)
        if self.drag_data['index'] is not None and current_index != self.drag_data['index']:
            if self.on_drag_reorder:
                self.on_drag_reorder(self.drag_data['index'], current_index)
            self.drag_data['index'] = current_index
    
    def _on_drop(self, event):
        """End drag operation."""
        self.drag_data['index'] = None
    
    # Public methods for controller to use
    
    def update_file_list(self, files: List[str]):
        """Update the file listbox with current files."""
        self.file_listbox.delete(0, tk.END)
        for i, f in enumerate(files, 1):
            display_name = f"{i}. {os.path.basename(f)}"
            self.file_listbox.insert(tk.END, display_name)
    
    def get_selected_index(self) -> Optional[int]:
        """Get the currently selected item index."""
        selection = self.file_listbox.curselection()
        return selection[0] if selection else None
    
    def set_selection(self, index: int):
        """Set the selection to a specific index."""
        self.file_listbox.selection_clear(0, tk.END)
        self.file_listbox.selection_set(index)
    
    def update_status(self, message: str, progress: Optional[float] = None):
        """Update status label and progress bar."""
        self.status_label.config(text=message)
        if progress is not None:
            self.progress_var.set(progress)
        self.root.update_idletasks()
    
    def set_convert_button_enabled(self, enabled: bool):
        """Enable or disable the convert buttons."""
        if enabled:
            self.convert_btn.config(state=tk.NORMAL, bg=self.colors['primary'])
            self.convert_btn.bind(
                '<Enter>', 
                lambda e: self.convert_btn.config(bg=self.colors['primary_hover'])
            )
            self.convert_btn.bind(
                '<Leave>', 
                lambda e: self.convert_btn.config(bg=self.colors['primary'])
            )
            self.convert_separate_btn.config(state=tk.NORMAL, bg=self.colors['success'])
            self.convert_separate_btn.bind(
                '<Enter>', 
                lambda e: self.convert_separate_btn.config(bg='#1e7b34')
            )
            self.convert_separate_btn.bind(
                '<Leave>', 
                lambda e: self.convert_separate_btn.config(bg=self.colors['success'])
            )
        else:
            self.convert_btn.config(state=tk.DISABLED, bg=self.colors['disabled'])
            self.convert_btn.unbind('<Enter>')
            self.convert_btn.unbind('<Leave>')
            self.convert_separate_btn.config(state=tk.DISABLED, bg=self.colors['disabled'])
            self.convert_separate_btn.unbind('<Enter>')
            self.convert_separate_btn.unbind('<Leave>')
    
    def show_warning(self, title: str, message: str):
        """Show a warning message box."""
        messagebox.showwarning(title, message)
    
    def show_error(self, title: str, message: str):
        """Show an error message box."""
        messagebox.showerror(title, message)
    
    def show_info(self, title: str, message: str):
        """Show an info message box."""
        messagebox.showinfo(title, message)
    
    def ask_save_file(self) -> Optional[str]:
        """Show save file dialog and return selected path."""
        return filedialog.asksaveasfilename(
            title="Save Merged PDF As",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            initialfile="merged_presentation.pdf"
        )
    
    def ask_open_files(self) -> List[str]:
        """Show open files dialog and return selected paths."""
        filetypes = [
            ("PowerPoint files", "*.pptx *.ppt"),
            ("PPTX files", "*.pptx"),
            ("PPT files", "*.ppt"),
            ("All files", "*.*")
        ]
        return list(filedialog.askopenfilenames(
            title="Select PowerPoint Files",
            filetypes=filetypes
        ))
    
    def ask_directory(self) -> Optional[str]:
        """Show directory selection dialog and return selected path."""
        return filedialog.askdirectory(
            title="Select Output Folder for PDFs"
        )
    
    def schedule(self, callback: Callable, *args):
        """Schedule a callback to run on the main thread."""
        self.root.after(0, callback, *args)
    
    @property
    def delete_temp_files(self) -> bool:
        """Get the delete temp files option value."""
        return self.delete_temp_var.get()
    
    @property
    def open_after_conversion(self) -> bool:
        """Get the open after conversion option value."""
        return self.open_after_var.get()
