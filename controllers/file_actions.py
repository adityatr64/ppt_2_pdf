import os
from typing import Any, Callable, Dict, List, Optional

from views import MainView


class FileActions:
    def __init__(
        self,
        view: MainView,
        get_active_task: Callable[[], Optional[Dict[str, Any]]],
        get_active_files: Callable[[], List[str]],
        refresh_list: Callable[[], None],
    ) -> None:
        self._view = view
        self._get_active_task = get_active_task
        self._get_active_files = get_active_files
        self._refresh_list = refresh_list

    def _can_edit_files(self) -> bool:
        active_task = self._get_active_task()
        if active_task and active_task["is_converting"]:
            self._view.show_warning("Task Running", "Cannot edit files while this task is converting.")
            return False
        return True

    def _can_reorder_files(self, warning_message: str) -> bool:
        active_task = self._get_active_task()
        if active_task and active_task["is_converting"]:
            self._view.show_warning("Task Running", warning_message)
            return False
        return True

    def add_files(self) -> None:
        if not self._can_edit_files():
            return

        new_files = self._view.ask_open_files()
        files = self._get_active_files()

        for file_path in new_files:
            if file_path not in files:
                files.append(file_path)

        self._refresh_list()

    def remove_selected(self) -> None:
        if not self._can_edit_files():
            return

        selected_index = self._view.get_selected_index()
        files = self._get_active_files()
        if selected_index is not None:
            files.pop(selected_index)
            self._refresh_list()

    def clear_all(self) -> None:
        if not self._can_edit_files():
            return

        active_task = self._get_active_task()
        if not active_task:
            return

        active_task["files"].clear()
        active_task["status"] = "Ready - Add files to get started"
        active_task["progress"] = 0.0
        self._refresh_list()

    def move_up(self) -> None:
        if not self._can_reorder_files("Cannot reorder files while this task is converting."):
            return

        selected_index = self._view.get_selected_index()
        files = self._get_active_files()
        if selected_index is not None and selected_index > 0:
            files[selected_index], files[selected_index - 1] = files[selected_index - 1], files[selected_index]
            self._refresh_list()
            self._view.set_selection(selected_index - 1)

    def move_down(self) -> None:
        if not self._can_reorder_files("Cannot reorder files while this task is converting."):
            return

        selected_index = self._view.get_selected_index()
        files = self._get_active_files()
        if selected_index is not None and selected_index < len(files) - 1:
            files[selected_index], files[selected_index + 1] = files[selected_index + 1], files[selected_index]
            self._refresh_list()
            self._view.set_selection(selected_index + 1)

    def drag_reorder(self, from_index: int, to_index: int) -> None:
        active_task = self._get_active_task()
        if active_task and active_task["is_converting"]:
            return

        files = self._get_active_files()
        if from_index is not None and 0 <= from_index < len(files):
            item = files.pop(from_index)
            files.insert(to_index, item)
            self._refresh_list()
            self._view.set_selection(to_index)

    def sort_files(self) -> None:
        if not self._can_reorder_files("Cannot sort files while this task is converting."):
            return

        files = self._get_active_files()
        files.sort(key=lambda path: os.path.basename(path).lower())
        self._refresh_list()
