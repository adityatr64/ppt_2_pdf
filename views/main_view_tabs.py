import tkinter as tk
from typing import Any, List, cast

import customtkinter as ctk

from .main_view_theme import set_button_cursor

# AI WARNING: This module contains AI-assisted code.
# Review and test changes carefully before release.


def clear_tab_widgets(view: Any) -> None:
    for row in view._tab_rows:
        try:
            row.destroy()
        except tk.TclError:
            pass
    view._tab_rows = []
    view._tab_buttons = {}


def build_tab_row(view: Any, tab_name: str, is_active: bool, is_running: bool) -> None:
    row = ctk.CTkFrame(view.task_tabs_container, fg_color="transparent")
    row.pack(side=tk.LEFT, padx=(0, 6))

    status_dot = "● " if is_running else ""
    tab_button = cast(Any, ctk.CTkButton(
        row,
        text=f"{status_dot}{tab_name}",
        command=lambda name=tab_name: view._on_task_tab_switched(name),
        width=120,
        height=30,
        font=view.font_regular,
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
        command=lambda name=tab_name: view._on_task_tab_closed(name),
        width=28,
        height=30,
        font=ctk.CTkFont(family=view.font_bold.cget("family"), size=13, weight="bold"),
    ))
    close_button.pack(side=tk.LEFT, padx=(2, 0))

    view._tab_rows.append(row)
    view._tab_buttons[tab_name] = (tab_button, close_button)
    set_button_cursor(tab_button, ["heart", "hand2", "arrow"])
    set_button_cursor(close_button, ["heart", "hand2", "arrow"])


def update_task_tabs(view: Any, tab_names: List[str], active_tab: str, running_tabs: List[str]) -> None:
    if not tab_names:
        tab_names = ["Task 1"]
        active_tab = "Task 1"

    view._updating_tabs = True
    clear_tab_widgets(view)
    for tab_name in tab_names:
        build_tab_row(view, tab_name, tab_name == active_tab, tab_name in running_tabs)
    view._updating_tabs = False
