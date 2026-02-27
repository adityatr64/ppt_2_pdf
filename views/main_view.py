"""
Main application view using CustomTkinter.

⚠️ THIS PART IS ENTIRELY MADE WITH AI ⚠️
⚠️ DON'T BLAME ME IF THE VIEW DIES I FUCKING HATE FRONTEND ⚠️

"""

import os
import sys
import tkinter as tk
import customtkinter as ctk
from tkinter import filedialog, messagebox
from typing import Any, Callable, List, Optional, cast

try:
    from .icons import apply_window_icon
except ImportError:
    try:
        from views.icons import apply_window_icon
    except ImportError:
        from icons import apply_window_icon


class MainView:
    """Main application window view."""

    def __init__(self, root: ctk.CTk):
        self.root = root
        self.icon_photos = []
        self.font_regular = ctk.CTkFont(size=12)
        self.font_bold = ctk.CTkFont(size=12, weight="bold")

        self.on_add_files: Optional[Callable[[], None]] = None
        self.on_remove_selected: Optional[Callable[[], None]] = None
        self.on_clear_all: Optional[Callable[[], None]] = None
        self.on_move_up: Optional[Callable[[], None]] = None
        self.on_move_down: Optional[Callable[[], None]] = None
        self.on_sort_files: Optional[Callable[[], None]] = None
        self.on_convert: Optional[Callable[[], None]] = None
        self.on_convert_separate: Optional[Callable[[], None]] = None
        self.on_drag_reorder: Optional[Callable[[int, int], None]] = None
        self.on_cancel: Optional[Callable[[], None]] = None
        self.on_new_task_tab: Optional[Callable[[], None]] = None
        self.on_close_task_tab: Optional[Callable[[str], None]] = None
        self.on_switch_task_tab: Optional[Callable[[str], None]] = None

        # Button references for enabling/disabling
        self.add_files_btn: Optional[Any] = None
        self.remove_btn: Optional[Any] = None
        self.clear_all_btn: Optional[Any] = None
        self.move_up_btn: Optional[Any] = None
        self.move_down_btn: Optional[Any] = None
        self.sort_btn: Optional[Any] = None
        self.convert_btn: Optional[Any] = None
        self.convert_separate_btn: Optional[Any] = None
        self.cancel_btn: Optional[Any] = None
        self.new_tab_btn: Optional[Any] = None

        # Queue status
        self.queue_label: Optional[Any] = None
        self._is_enabled = True

        self.drag_data = {"index": None}
        self._hover_index: Optional[int] = None
        self._updating_tabs = False
        self._tab_rows: List[Any] = []
        self._tab_buttons: dict = {}

        self._setup_window()
        self._load_app_fonts()
        self._setup_ui()

    def _resource_path(self, relative_path: str) -> str:
        base_path = getattr(sys, "_MEIPASS", os.path.abspath("."))
        return os.path.join(base_path, relative_path)

    def _load_app_fonts(self):
        regular_font = self._resource_path(os.path.join("assets", "fonts", "TTF", "VictorMono-ExtraLight.ttf"))
        bold_font = self._resource_path(os.path.join("assets", "fonts", "TTF", "VictorMono-SemiBold.ttf"))

        try:
            ctk.FontManager.load_font(regular_font)
            ctk.FontManager.load_font(bold_font)
            self.font_regular = ctk.CTkFont(family="Victor Mono", size=12)
            self.font_bold = ctk.CTkFont(family="Victor Mono", size=12, weight="bold")
        except Exception:
            self.font_regular = ctk.CTkFont(size=12)
            self.font_bold = ctk.CTkFont(size=12, weight="bold")

    def _apply_listbox_theme(self):
        is_dark = ctk.get_appearance_mode().lower() == "dark"
        if is_dark:
            self.file_listbox.configure(
                bg="#1f1f1f",
                fg="#f5f5f5",
                selectbackground="#3b82f6",
                selectforeground="#ffffff",
            )
        else:
            self.file_listbox.configure(
                bg="#ffffff",
                fg="#202124",
                selectbackground="#3b82f6",
                selectforeground="#ffffff",
            )

    def _setup_window(self):
        self.root.title("PPT 2 PDF")
        self.root.geometry("900x700")
        self.root.minsize(760, 620)
        self.icon_photos = apply_window_icon(self.root)

    def _setup_ui(self):
        container = ctk.CTkScrollableFrame(self.root, corner_radius=0)
        container.pack(fill=tk.BOTH, expand=True, padx=16, pady=16)

        title = ctk.CTkLabel(
            container,
            text="PowerPoint to PDF Converter & Merger",
            font=ctk.CTkFont(family=self.font_bold.cget("family"), size=22, weight="bold"),
        )
        title.pack(anchor="w", pady=(4, 12), padx=10)

        self._setup_file_list(container)
        self._setup_buttons(container)
        self._setup_options(container)
        self._setup_progress(container)
        self._setup_convert_buttons(container)

    def _setup_file_list(self, parent):
        list_frame = ctk.CTkFrame(parent)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        header = ctk.CTkLabel(
            list_frame,
            text="Selected Files",
            font=ctk.CTkFont(family=self.font_bold.cget("family"), size=15, weight="bold"),
        )
        header.pack(anchor="w", padx=12, pady=(10, 2))

        hint = ctk.CTkLabel(
            list_frame,
            text="Drag to reorder or use controls below",
            font=self.font_regular,
            text_color=("#5f6368", "#9aa0a6"),
        )
        hint.pack(anchor="w", padx=12, pady=(0, 8))

        tabs_row = ctk.CTkFrame(list_frame, fg_color="transparent")
        tabs_row.pack(fill=tk.X, padx=10, pady=(0, 8))

        self.task_tabs_container = ctk.CTkFrame(tabs_row, fg_color="transparent")
        self.task_tabs_container.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))

        new_tab_btn = cast(Any, ctk.CTkButton(
            tabs_row,
            text="+ New Task",
            width=110,
            command=lambda: self._invoke(self.on_new_task_tab),
            font=self.font_regular,
        ))
        new_tab_btn.pack(side=tk.LEFT, padx=(0, 6))
        self.new_tab_btn = new_tab_btn

        list_container = ctk.CTkFrame(list_frame)
        list_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        scrollbar = ctk.CTkScrollbar(list_container, orientation="vertical")
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y, pady=8, padx=(0, 8))

        self.file_listbox = tk.Listbox(
            list_container,
            selectmode=tk.SINGLE,
            yscrollcommand=scrollbar.set,
            font=(self.font_regular.cget("family"), 11),
            activestyle="dotbox",
            borderwidth=0,
            highlightthickness=0,
            relief="flat",
            bg="#ffffff",
            fg="#202124",
            selectbackground="#3b82f6",
            selectforeground="#ffffff",
            cursor="hand1",
        )
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=8, pady=8)
        scrollbar.configure(command=self.file_listbox.yview)
        self.file_listbox.configure(yscrollcommand=scrollbar.set)
        self._apply_listbox_theme()

        self.file_listbox.bind("<Button-1>", self._on_click)
        self.file_listbox.bind("<B1-Motion>", self._on_drag)
        self.file_listbox.bind("<ButtonRelease-1>", self._on_drop)
        self.file_listbox.bind("<Motion>", self._on_list_hover)
        self.file_listbox.bind("<Leave>", lambda _event: self._set_list_cursor(["arrow"]))
        self.file_listbox.bind("<MouseWheel>", self._on_list_mousewheel)
        self.file_listbox.bind("<Button-4>", self._on_list_mousewheel)
        self.file_listbox.bind("<Button-5>", self._on_list_mousewheel)

    def _set_list_cursor(self, candidates: List[str]):
        for cursor_name in candidates:
            try:
                self.file_listbox.configure(cursor=cursor_name)
                return
            except tk.TclError:
                continue

    def _set_button_cursor(self, button: Optional[Any], candidates: List[str]):
        if button is None:
            return
        for cursor_name in candidates:
            try:
                button.configure(cursor=cursor_name)
                return
            except tk.TclError:
                continue

    def _on_task_tab_closed(self, tab_name: str):
        if self.on_close_task_tab:
            self.on_close_task_tab(tab_name)

    def _on_task_tab_switched(self, tab_name: str):
        if self._updating_tabs:
            return
        if self.on_switch_task_tab:
            self.on_switch_task_tab(tab_name)

    def _clear_tab_widgets(self):
        for row in self._tab_rows:
            try:
                row.destroy()
            except tk.TclError:
                pass
        self._tab_rows = []
        self._tab_buttons = {}

    def _build_tab_row(self, tab_name: str, is_active: bool, is_running: bool):
        row = ctk.CTkFrame(self.task_tabs_container, fg_color="transparent")
        row.pack(side=tk.LEFT, padx=(0, 6))

        status_dot = "● " if is_running else ""
        tab_button = cast(Any, ctk.CTkButton(
            row,
            text=f"{status_dot}{tab_name}",
            command=lambda name=tab_name: self._on_task_tab_switched(name),
            width=120,
            height=30,
            font=self.font_regular,
        ))
        tab_button.pack(side=tk.LEFT)

        if is_active:
            tab_button.configure(
                fg_color=("#1f6aa5", "#1f6aa5"),
                hover_color=("#185786", "#185786"),
            )

        close_button = cast(Any, ctk.CTkButton(
            row,
            text="×",
            command=lambda name=tab_name: self._on_task_tab_closed(name),
            width=28,
            height=30,
            font=ctk.CTkFont(family=self.font_bold.cget("family"), size=13, weight="bold"),
        ))
        close_button.pack(side=tk.LEFT, padx=(2, 0))

        self._tab_rows.append(row)
        self._tab_buttons[tab_name] = (tab_button, close_button)
        self._set_button_cursor(tab_button, ["heart", "hand2", "arrow"])
        self._set_button_cursor(close_button, ["heart", "hand2", "arrow"])

    def _set_hover_row(self, index: Optional[int]):
        if self._hover_index == index:
            return

        is_dark = ctk.get_appearance_mode().lower() == "dark"
        normal_bg = "#1f1f1f" if is_dark else "#ffffff"
        hover_bg = "#2a2a2a" if is_dark else "#eef4ff"

        if self._hover_index is not None and 0 <= self._hover_index < self.file_listbox.size():
            try:
                self.file_listbox.itemconfig(self._hover_index, background=normal_bg)
            except tk.TclError:
                pass

        self._hover_index = index

        if self._hover_index is not None and 0 <= self._hover_index < self.file_listbox.size():
            self.file_listbox.activate(self._hover_index)
            try:
                self.file_listbox.itemconfig(self._hover_index, background=hover_bg)
            except tk.TclError:
                pass

    def _is_pointer_on_item(self, event) -> bool:
        if self.file_listbox.size() == 0:
            return False
        index = self.file_listbox.nearest(event.y)
        if index < 0 or index >= self.file_listbox.size():
            return False
        bbox = self.file_listbox.bbox(index)
        if not bbox:
            return False
        _, y, _, height = bbox
        return y <= event.y <= y + height

    def _on_list_hover(self, event):
        if self.drag_data["index"] is not None:
            return
        if self._is_pointer_on_item(event):
            hovered_index = self.file_listbox.nearest(event.y)
            self._set_hover_row(hovered_index)
            self._set_list_cursor(["openhand", "hand1", "pointinghand", "arrow"])
        else:
            self._set_hover_row(None)
            self._set_list_cursor(["arrow"])

    def _on_list_mousewheel(self, event):
        first, last = self.file_listbox.yview()

        if hasattr(event, "num") and event.num in (4, 5):
            delta = 1 if event.num == 4 else -1
        else:
            delta = 1 if event.delta > 0 else -1

        scrolling_up = delta > 0
        at_top = first <= 0.0
        at_bottom = last >= 1.0

        if (scrolling_up and at_top) or ((not scrolling_up) and at_bottom):
            return None

        units = -1 if scrolling_up else 1
        self.file_listbox.yview_scroll(units, "units")
        return "break"

    def _setup_buttons(self, parent):
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill=tk.X, padx=10, pady=(0, 10))

        left = ctk.CTkFrame(row, fg_color="transparent")
        left.pack(side=tk.LEFT)

        add_files_btn = cast(Any, ctk.CTkButton(left, text="+ Add Files", command=lambda: self._invoke(self.on_add_files), width=110, font=self.font_regular))
        add_files_btn.pack(side=tk.LEFT, padx=4)
        self.add_files_btn = add_files_btn
        
        remove_btn = cast(Any, ctk.CTkButton(left, text="Remove", command=lambda: self._invoke(self.on_remove_selected), width=95, font=self.font_regular))
        remove_btn.pack(side=tk.LEFT, padx=4)
        self.remove_btn = remove_btn
        
        clear_all_btn = cast(Any, ctk.CTkButton(left, text="Clear All", command=lambda: self._invoke(self.on_clear_all), width=95, font=self.font_regular))
        clear_all_btn.pack(side=tk.LEFT, padx=4)
        self.clear_all_btn = clear_all_btn

        right = ctk.CTkFrame(row, fg_color="transparent")
        right.pack(side=tk.RIGHT)

        move_up_btn = cast(Any, ctk.CTkButton(right, text="Move Up", command=lambda: self._invoke(self.on_move_up), width=95, font=self.font_regular))
        move_up_btn.pack(side=tk.LEFT, padx=4)
        self.move_up_btn = move_up_btn
        
        move_down_btn = cast(Any, ctk.CTkButton(right, text="Move Down", command=lambda: self._invoke(self.on_move_down), width=95, font=self.font_regular))
        move_down_btn.pack(side=tk.LEFT, padx=4)
        self.move_down_btn = move_down_btn
        
        sort_btn = cast(Any, ctk.CTkButton(right, text="Sort", command=lambda: self._invoke(self.on_sort_files), width=85, font=self.font_regular))
        sort_btn.pack(side=tk.LEFT, padx=4)
        self.sort_btn = sort_btn

    def _setup_options(self, parent):
        options = ctk.CTkFrame(parent)
        options.pack(fill=tk.X, padx=10, pady=(0, 10))

        ctk.CTkLabel(
            options,
            text="Options",
            font=ctk.CTkFont(family=self.font_bold.cget("family"), size=14, weight="bold"),
        ).pack(anchor="w", padx=12, pady=(10, 4))

        ctk.CTkLabel(
            options,
            text="Temporary PDF files are always deleted after merge",
            font=self.font_regular,
            text_color=("#5f6368", "#9aa0a6"),
        ).pack(anchor="w", padx=12, pady=(0, 6))

        self.open_after_var = tk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            options,
            text="Open output after conversion",
            variable=self.open_after_var,
            font=self.font_regular,
        ).pack(anchor="w", padx=12, pady=(0, 10))

    def _setup_progress(self, parent):
        progress_frame = ctk.CTkFrame(parent)
        progress_frame.pack(fill=tk.X, padx=10, pady=(0, 10))

        self.status_label = ctk.CTkLabel(
            progress_frame,
            text="Ready - Add files to get started",
            font=self.font_regular,
            anchor="w",
        )
        self.status_label.pack(fill=tk.X, padx=12, pady=(10, 6))

        queue_label = cast(Any, ctk.CTkLabel(
            progress_frame,
            text="",
            font=self.font_regular,
            text_color=("#5f6368", "#9aa0a6"),
            anchor="w",
        ))
        queue_label.pack(fill=tk.X, padx=12, pady=(0, 4))
        self.queue_label = queue_label

        self.progress_var = tk.DoubleVar(value=0.0)
        self.progress_bar = ctk.CTkProgressBar(progress_frame)
        self.progress_bar.pack(fill=tk.X, padx=12, pady=(0, 12))
        self.progress_bar.set(0)

    def _setup_convert_buttons(self, parent):
        btn_frame = ctk.CTkFrame(parent, fg_color="transparent")
        btn_frame.pack(pady=(4, 2))

        convert_btn = cast(Any, ctk.CTkButton(
            btn_frame,
            text="Convert & Merge to PDF",
            command=lambda: self._invoke(self.on_convert),
            width=220,
            height=40,
            font=ctk.CTkFont(family=self.font_bold.cget("family"), size=14, weight="bold"),
        ))
        convert_btn.pack(side=tk.LEFT, padx=6)
        self.convert_btn = convert_btn

        convert_separate_btn = cast(Any, ctk.CTkButton(
            btn_frame,
            text="Make Separate PDFs",
            command=lambda: self._invoke(self.on_convert_separate),
            width=200,
            height=40,
            font=ctk.CTkFont(family=self.font_bold.cget("family"), size=14, weight="bold"),
        ))
        convert_separate_btn.pack(side=tk.LEFT, padx=6)
        self.convert_separate_btn = convert_separate_btn

        cancel_btn = cast(Any, ctk.CTkButton(
            btn_frame,
            text="Cancel",
            command=lambda: self._invoke(self.on_cancel),
            width=120,
            height=40,
            font=ctk.CTkFont(family=self.font_bold.cget("family"), size=14, weight="bold"),
            fg_color="#ff6b6b",
            hover_color="#ff5252",
        ))
        cancel_btn.pack(side=tk.LEFT, padx=6)
        cancel_btn.pack_forget()
        self.cancel_btn = cancel_btn

        self._set_button_cursor(self.convert_btn, ["heart", "hand2", "arrow"])
        self._set_button_cursor(self.convert_separate_btn, ["heart", "hand2", "arrow"])
        self._set_button_cursor(self.cancel_btn, ["heart", "hand2", "arrow"])
        self._set_button_cursor(self.new_tab_btn, ["heart", "hand2", "arrow"])

    def _invoke(self, callback):
        if callback:
            callback()

    def _on_click(self, event):
        if self._is_pointer_on_item(event):
            index = self.file_listbox.nearest(event.y)
            self.drag_data["index"] = index
            self.file_listbox.selection_clear(0, tk.END)
            self.file_listbox.selection_set(index)
            self.file_listbox.activate(index)
            self._set_hover_row(index)
            self._set_list_cursor(["closedhand", "exchange", "sizeall", "hand1"])
        else:
            self.drag_data["index"] = None
            self._set_hover_row(None)
            self._set_list_cursor(["arrow"])

    def _on_drag(self, event):
        current_index = self.file_listbox.nearest(event.y)
        self._set_list_cursor(["closedhand", "exchange", "sizeall", "hand1"])
        if self.drag_data["index"] is not None and current_index != self.drag_data["index"]:
            if self.on_drag_reorder:
                self.on_drag_reorder(self.drag_data["index"], current_index)
            self.drag_data["index"] = current_index
            self._set_hover_row(current_index)

    def _on_drop(self, event):
        self.drag_data["index"] = None
        if self._is_pointer_on_item(event):
            index = self.file_listbox.nearest(event.y)
            self._set_hover_row(index)
            self._set_list_cursor(["openhand", "hand1", "pointinghand", "arrow"])
        else:
            self._set_hover_row(None)
            self._set_list_cursor(["arrow"])

    def update_file_list(self, files: List[str]):
        self.file_listbox.delete(0, tk.END)
        for i, file_path in enumerate(files, 1):
            self.file_listbox.insert(tk.END, f"{i}. {os.path.basename(file_path)}")
        self._hover_index = None

    def update_task_tabs(self, tab_names: List[str], active_tab: str, running_tabs: List[str]):
        if not tab_names:
            tab_names = ["Task 1"]
            active_tab = "Task 1"

        self._updating_tabs = True
        self._clear_tab_widgets()
        for tab_name in tab_names:
            self._build_tab_row(tab_name, tab_name == active_tab, tab_name in running_tabs)
        self._updating_tabs = False

    def update_task_actions(self, is_running: bool):
        if self.convert_btn is None or self.convert_separate_btn is None or self.cancel_btn is None:
            return

        if is_running:
            self.convert_btn.configure(state="disabled")
            self.convert_separate_btn.configure(state="disabled")
            self.cancel_btn.pack(side=tk.LEFT, padx=6)
            self.cancel_btn.configure(state="normal")
            self._set_button_cursor(self.cancel_btn, ["heart", "hand2", "arrow"])
        else:
            self.convert_btn.configure(state="normal")
            self.convert_separate_btn.configure(state="normal")
            self.cancel_btn.pack_forget()

    def get_selected_index(self) -> Optional[int]:
        selection = self.file_listbox.curselection()
        return selection[0] if selection else None

    def set_selection(self, index: int):
        self.file_listbox.selection_clear(0, tk.END)
        self.file_listbox.selection_set(index)

    def update_status(self, message: str, progress: Optional[float] = None):
        self.status_label.configure(text=message)
        if progress is not None:
            self.progress_var.set(progress)
            self.progress_bar.set(max(0.0, min(1.0, progress / 100.0)))
        self.root.update_idletasks()

    def set_ui_enabled(self, enabled: bool):
        """Generic method to enable/disable all interactive UI elements during conversion."""
        self._is_enabled = enabled
        state = "normal" if enabled else "disabled"
        cursor = ["heart", "hand2", "arrow"] if enabled else ["arrow"]

        # Disable file management buttons
        file_buttons = [
            self.add_files_btn,
            self.remove_btn,
            self.clear_all_btn,
            self.move_up_btn,
            self.move_down_btn,
            self.sort_btn,
            self.new_tab_btn,
        ]
        for btn in file_buttons:
            if btn is not None:
                btn.configure(state=state)

        # Disable/enable convert buttons (hide during conversion)
        if enabled:
            if self.convert_btn is not None:
                self.convert_btn.pack(side=tk.LEFT, padx=6)
            if self.convert_separate_btn is not None:
                self.convert_separate_btn.pack(side=tk.LEFT, padx=6)
            if self.cancel_btn is not None:
                self.cancel_btn.pack(side=tk.LEFT, padx=6)
            self._set_button_cursor(self.convert_btn, cursor)
            self._set_button_cursor(self.convert_separate_btn, cursor)
            self._set_button_cursor(self.cancel_btn, cursor)
        else:
            if self.convert_btn is not None:
                self.convert_btn.configure(state="disabled")
            if self.convert_separate_btn is not None:
                self.convert_separate_btn.configure(state="disabled")
            if self.cancel_btn is not None:
                self.cancel_btn.configure(state="normal")
            self._set_button_cursor(self.cancel_btn, cursor)

        for btn in file_buttons:
            self._set_button_cursor(btn, cursor)

    def update_queue_status(self, queue_count: int):
        """Update running task count label."""
        if self.queue_label is None:
            return
        if queue_count > 0:
            self.queue_label.configure(text=f"⚙️ {queue_count} task(s) running")
        else:
            self.queue_label.configure(text="No running tasks")

    def show_warning(self, title: str, message: str):
        messagebox.showwarning(title, message)

    def show_error(self, title: str, message: str):
        messagebox.showerror(title, message)

    def show_info(self, title: str, message: str):
        messagebox.showinfo(title, message)

    def ask_confirm(self, title: str, message: str) -> bool:
        return messagebox.askyesno(title, message)

    def ask_save_file(self, initialfile: str = "merged_presentation.pdf") -> Optional[str]:
        return filedialog.asksaveasfilename(
            title="Save Merged PDF As",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            initialfile=initialfile,
        )

    def ask_open_files(self) -> List[str]:
        filetypes = [
            ("PowerPoint files", "*.pptx *.ppt"),
            ("PPTX files", "*.pptx"),
            ("PPT files", "*.ppt"),
            ("All files", "*.*"),
        ]
        return list(
            filedialog.askopenfilenames(
                title="Select PowerPoint Files",
                filetypes=filetypes,
            )
        )

    def ask_directory(self) -> Optional[str]:
        return filedialog.askdirectory(title="Select Output Folder for PDFs")

    def schedule(self, callback: Callable, *args):
        self.root.after(0, callback, *args)

    @property
    def delete_temp_files(self) -> bool:
        return True

    @property
    def open_after_conversion(self) -> bool:
        return self.open_after_var.get()
