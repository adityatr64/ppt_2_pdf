import tkinter as tk
from typing import Any

import customtkinter as ctk

from .main_view_theme import set_list_cursor

# AI WARNING: This module contains AI-assisted code.
# Review and test changes carefully before release.


def set_hover_row(view: Any, index):
    if view._hover_index == index:
        return

    is_dark = ctk.get_appearance_mode().lower() == "dark"
    normal_bg = "#1f1f1f" if is_dark else "#ffffff"
    hover_bg = "#2a2a2a" if is_dark else "#eef4ff"

    if view._hover_index is not None and 0 <= view._hover_index < view.file_listbox.size():
        try:
            view.file_listbox.itemconfig(view._hover_index, background=normal_bg)
        except tk.TclError:
            pass

    view._hover_index = index

    if view._hover_index is not None and 0 <= view._hover_index < view.file_listbox.size():
        view.file_listbox.activate(view._hover_index)
        try:
            view.file_listbox.itemconfig(view._hover_index, background=hover_bg)
        except tk.TclError:
            pass


def is_pointer_on_item(view: Any, event) -> bool:
    if view.file_listbox.size() == 0:
        return False
    index = view.file_listbox.nearest(event.y)
    if index < 0 or index >= view.file_listbox.size():
        return False
    bbox = view.file_listbox.bbox(index)
    if not bbox:
        return False
    _, y, _, height = bbox
    return y <= event.y <= y + height


def on_list_hover(view: Any, event) -> None:
    if view.drag_data["index"] is not None:
        return
    if is_pointer_on_item(view, event):
        hovered_index = view.file_listbox.nearest(event.y)
        set_hover_row(view, hovered_index)
        set_list_cursor(view, ["openhand", "hand1", "pointinghand", "arrow"])
    else:
        set_hover_row(view, None)
        set_list_cursor(view, ["arrow"])


def on_list_mousewheel(view: Any, event):
    first, last = view.file_listbox.yview()

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
    view.file_listbox.yview_scroll(units, "units")
    return "break"


def on_click(view: Any, event) -> None:
    if is_pointer_on_item(view, event):
        index = view.file_listbox.nearest(event.y)
        view.drag_data["index"] = index
        view.file_listbox.selection_clear(0, tk.END)
        view.file_listbox.selection_set(index)
        view.file_listbox.activate(index)
        set_hover_row(view, index)
        set_list_cursor(view, ["closedhand", "exchange", "sizeall", "hand1"])
    else:
        view.drag_data["index"] = None
        set_hover_row(view, None)
        set_list_cursor(view, ["arrow"])


def on_drag(view: Any, event) -> None:
    if view.drag_data["index"] is None:
        return

    edge_margin = 24
    list_height = view.file_listbox.winfo_height()
    if event.y < edge_margin:
        view.file_listbox.yview_scroll(-1, "units")
    elif event.y > max(edge_margin, list_height - edge_margin):
        view.file_listbox.yview_scroll(1, "units")

    current_index = view.file_listbox.nearest(event.y)
    set_list_cursor(view, ["closedhand", "exchange", "sizeall", "hand1"])
    if current_index != view.drag_data["index"]:
        if view.on_drag_reorder:
            view.on_drag_reorder(view.drag_data["index"], current_index)
        view.drag_data["index"] = current_index
        set_hover_row(view, current_index)


def on_drop(view: Any, event) -> None:
    view.drag_data["index"] = None
    if is_pointer_on_item(view, event):
        index = view.file_listbox.nearest(event.y)
        set_hover_row(view, index)
        set_list_cursor(view, ["openhand", "hand1", "pointinghand", "arrow"])
    else:
        set_hover_row(view, None)
        set_list_cursor(view, ["arrow"])
