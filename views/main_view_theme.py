import os
import sys
import tkinter as tk
from typing import Any, List, Optional

import customtkinter as ctk

# AI WARNING: This module contains AI-assisted code.
# Review and test changes carefully before release.


def resource_path(relative_path: str) -> str:
    base_path = getattr(sys, "_MEIPASS", os.path.abspath("."))
    return os.path.join(base_path, relative_path)


def load_app_fonts(view: Any) -> None:
    regular_font = resource_path(os.path.join("assets", "fonts", "TTF", "VictorMono-ExtraLight.ttf"))
    bold_font = resource_path(os.path.join("assets", "fonts", "TTF", "VictorMono-SemiBold.ttf"))

    try:
        ctk.FontManager.load_font(regular_font)
        ctk.FontManager.load_font(bold_font)
        view.font_regular = ctk.CTkFont(family="Victor Mono", size=12)
        view.font_bold = ctk.CTkFont(family="Victor Mono", size=12, weight="bold")
    except Exception:
        view.font_regular = ctk.CTkFont(size=12)
        view.font_bold = ctk.CTkFont(size=12, weight="bold")


def apply_listbox_theme(view: Any) -> None:
    is_dark = ctk.get_appearance_mode().lower() == "dark"
    if is_dark:
        view.file_listbox.configure(
            bg="#1f1f1f",
            fg="#f5f5f5",
            selectbackground="#3b82f6",
            selectforeground="#ffffff",
        )
    else:
        view.file_listbox.configure(
            bg="#ffffff",
            fg="#202124",
            selectbackground="#3b82f6",
            selectforeground="#ffffff",
        )


def set_list_cursor(view: Any, candidates: List[str]) -> None:
    for cursor_name in candidates:
        try:
            view.file_listbox.configure(cursor=cursor_name)
            return
        except tk.TclError:
            continue


def set_button_cursor(button: Optional[Any], candidates: List[str]) -> None:
    if button is None:
        return
    for cursor_name in candidates:
        try:
            button.configure(cursor=cursor_name)
            return
        except tk.TclError:
            continue
