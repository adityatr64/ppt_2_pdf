"""
Main application view using CustomTkinter.

⚠️ THIS PART IS ENTIRELY MADE WITH AI ⚠️
⚠️ DON'T BLAME ME IF THE VIEW DIES I FUCKING HATE FRONTEND ⚠️

"""

import os
import tkinter as tk
from typing import Any, Callable, List, Optional, cast

import customtkinter as ctk

from .icons import apply_window_icon
from .main_view_dialogs import (
    ask_confirm,
    ask_directory,
    ask_open_files,
    ask_save_file,
    show_error,
    show_info,
    show_warning,
)
from .main_view_interactions import (
    on_click,
    on_drag,
    on_drop,
    on_list_hover,
    on_list_mousewheel,
)
from .main_view_tabs import update_task_tabs
from .main_view_theme import apply_listbox_theme, load_app_fonts, set_button_cursor, set_list_cursor


"""AI-assisted code notice: parts of this UI were generated with AI and should be looked at with a grain of salt :3 """


class MainView:
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

        self.queue_label: Optional[Any] = None

        self.drag_data = {"index": None}
        self._hover_index: Optional[int] = None
        self._updating_tabs = False
        self._tab_rows: List[Any] = []
        self._tab_buttons: dict = {}

        self._setup_window()
        load_app_fonts(self)
        self._setup_ui()

    def _setup_window(self) -> None:
        self.root.title("PPT 2 PDF")
        self.root.geometry("900x700")
        self.root.minsize(760, 620)
        self.icon_photos = apply_window_icon(self.root)

    def _setup_ui(self) -> None:
        container = ctk.CTkFrame(self.root, corner_radius=0)
        container.pack(fill=tk.BOTH, expand=True, padx=16, pady=16)

        title = ctk.CTkLabel(
            container,
            text="PowerPoint to PDF Converter & Merger",
            font=ctk.CTkFont(family=self.font_bold.cget("family"), size=22, weight="bold"),
        )
        title.pack(anchor="w", pady=(4, 12), padx=10)

        is_dark = ctk.get_appearance_mode().lower() == "dark"
        pane_bg = "#3f3f3f" if is_dark else "#d7d7d7"
        content_pane = tk.PanedWindow(
            container,
            orient=tk.VERTICAL,
            sashrelief=tk.FLAT,
            sashwidth=10,
            sashpad=0,
            showhandle=False,
            opaqueresize=True,
            bd=0,
            bg=pane_bg,
            proxybackground=pane_bg,
            proxyborderwidth=0,
            sashcursor="sb_v_double_arrow",
        )
        content_pane.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 4))

        list_section = ctk.CTkFrame(content_pane)
        controls_pane = tk.Frame(content_pane, bd=0, highlightthickness=0, bg=pane_bg)
        controls_section = ctk.CTkScrollableFrame(controls_pane, corner_radius=6)
        controls_section.pack(fill=tk.BOTH, expand=True)

        content_pane.add(list_section, minsize=220)
        content_pane.add(controls_pane, minsize=220)

        self._setup_file_list(list_section)
        self._setup_buttons(controls_section)
        self._setup_options(controls_section)
        self._setup_progress(controls_section)
        self._setup_convert_buttons(controls_section)

    def _setup_file_list(self, parent) -> None:
        list_frame = ctk.CTkFrame(parent)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=0, pady=0)

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
            font=(self.font_regular.cget("family"), int(self.font_regular.cget("size"))),
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
        apply_listbox_theme(self)

        self.file_listbox.bind("<Button-1>", self._on_click)
        self.file_listbox.bind("<B1-Motion>", self._on_drag)
        self.file_listbox.bind("<ButtonRelease-1>", self._on_drop)
        self.file_listbox.bind("<Motion>", self._on_list_hover)
        self.file_listbox.bind("<Leave>", lambda _event: set_list_cursor(self, ["arrow"]))
        self.file_listbox.bind("<MouseWheel>", self._on_list_mousewheel)
        self.file_listbox.bind("<Button-4>", self._on_list_mousewheel)
        self.file_listbox.bind("<Button-5>", self._on_list_mousewheel)

    def _setup_buttons(self, parent) -> None:
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

    def _setup_options(self, parent) -> None:
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

    def _setup_progress(self, parent) -> None:
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

    def _setup_convert_buttons(self, parent) -> None:
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

        set_button_cursor(self.convert_btn, ["heart", "hand2", "arrow"])
        set_button_cursor(self.convert_separate_btn, ["heart", "hand2", "arrow"])
        set_button_cursor(self.cancel_btn, ["heart", "hand2", "arrow"])
        set_button_cursor(self.new_tab_btn, ["heart", "hand2", "arrow"])

    def _invoke(self, callback) -> None:
        if callback:
            callback()

    def _on_task_tab_closed(self, tab_name: str) -> None:
        if self.on_close_task_tab:
            self.on_close_task_tab(tab_name)

    def _on_task_tab_switched(self, tab_name: str) -> None:
        if self._updating_tabs:
            return
        if self.on_switch_task_tab:
            self.on_switch_task_tab(tab_name)

    def _on_list_hover(self, event) -> None:
        on_list_hover(self, event)

    def _on_list_mousewheel(self, event):
        return on_list_mousewheel(self, event)

    def _on_click(self, event) -> None:
        on_click(self, event)

    def _on_drag(self, event) -> None:
        on_drag(self, event)

    def _on_drop(self, event) -> None:
        on_drop(self, event)

    def update_file_list(self, files: List[str]) -> None:
        first_visible, _ = self.file_listbox.yview()
        selected_index = self.get_selected_index()

        self.file_listbox.delete(0, tk.END)
        for i, file_path in enumerate(files, 1):
            self.file_listbox.insert(tk.END, f"{i}. {os.path.basename(file_path)}")

        if files:
            preferred_index = self.drag_data["index"] if self.drag_data["index"] is not None else selected_index
            if preferred_index is not None:
                preferred_index = max(0, min(preferred_index, len(files) - 1))
                self.file_listbox.selection_set(preferred_index)
                self.file_listbox.activate(preferred_index)

            self.file_listbox.yview_moveto(first_visible)

        self._hover_index = None

    def update_task_tabs(self, tab_names: List[str], active_tab: str, running_tabs: List[str]) -> None:
        update_task_tabs(self, tab_names, active_tab, running_tabs)

    def update_task_actions(self, is_running: bool) -> None:
        if self.convert_btn is None or self.convert_separate_btn is None or self.cancel_btn is None:
            return

        if is_running:
            self.convert_btn.configure(state="disabled")
            self.convert_separate_btn.configure(state="disabled")
            self.cancel_btn.pack(side=tk.LEFT, padx=6)
            self.cancel_btn.configure(state="normal")
            set_button_cursor(self.cancel_btn, ["heart", "hand2", "arrow"])
        else:
            self.convert_btn.configure(state="normal")
            self.convert_separate_btn.configure(state="normal")
            self.cancel_btn.pack_forget()

    def get_selected_index(self) -> Optional[int]:
        selection = self.file_listbox.curselection()
        return selection[0] if selection else None

    def set_selection(self, index: int) -> None:
        self.file_listbox.selection_clear(0, tk.END)
        self.file_listbox.selection_set(index)

    def update_status(self, message: str, progress: Optional[float] = None) -> None:
        self.status_label.configure(text=message)
        if progress is not None:
            self.progress_var.set(progress)
            self.progress_bar.set(max(0.0, min(1.0, progress / 100.0)))
        self.root.update_idletasks()

    def update_queue_status(self, queue_count: int) -> None:
        if self.queue_label is None:
            return
        if queue_count > 0:
            self.queue_label.configure(text=f"⚙️ {queue_count} task(s) running")
        else:
            self.queue_label.configure(text="No running tasks")

    def show_warning(self, title: str, message: str) -> None:
        show_warning(title, message)

    def show_error(self, title: str, message: str) -> None:
        show_error(title, message)

    def show_info(self, title: str, message: str) -> None:
        show_info(title, message)

    def ask_confirm(self, title: str, message: str) -> bool:
        return ask_confirm(title, message)

    def ask_save_file(self, initialfile: str = "merged_presentation.pdf") -> Optional[str]:
        return ask_save_file(initialfile)

    def ask_open_files(self) -> List[str]:
        return ask_open_files()

    def ask_directory(self) -> Optional[str]:
        return ask_directory()

    def schedule(self, callback: Callable, *args) -> None:
        self.root.after(0, callback, *args)

    @property
    def delete_temp_files(self) -> bool:
        return True

    @property
    def open_after_conversion(self) -> bool:
        return self.open_after_var.get()
