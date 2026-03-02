from typing import Any, Dict, List, Optional

from views import MainView
from services import ConversionService

from .conversion_actions import ConversionActions
from .file_actions import FileActions
from .system_ops import open_path
from .tab_display import compute_tab_display_mapping
from .task_manager import TaskManager

# AI WARNING: This module contains AI-assisted code.


class AppController:
    MAX_TAB_LABEL_LENGTH = 18

    def __init__(self, view: MainView):
        self.view = view
        self.service = ConversionService()
        self.task_manager = TaskManager()
        self._display_to_internal_task_name: Dict[str, str] = {}

        self.file_actions = FileActions(
            view=self.view,
            get_active_task=self._active_task,
            get_active_files=self._active_files,
            refresh_list=self._refresh_list,
        )
        self.conversion_actions = ConversionActions(
            view=self.view,
            find_task=self._find_task,
            is_task_cancelled=self.task_manager.is_task_cancelled,
            refresh_list=self._refresh_list,
            finalize_successful_task=self._finalize_successful_task,
            open_path=open_path,
        )

        self.task_manager.create_task()
        self._bind_callbacks()
        self._refresh_list()
        self._check_backend_availability()

    def _check_backend_availability(self) -> None:
        try:
            available = self.service.get_available_backends()
            if not available:
                self.view.show_warning("No Backend Found", self.service.get_install_message())
                self.view.update_status("No backend found - install required app", 0)
                return

            backend_name = self.service.get_active_backend_name()
            self.view.update_status(f"Ready - Backend: {backend_name}", 0)
        except Exception as exc:
            self.view.show_warning("Backend Check Failed", str(exc))
            self.view.update_status("Backend check failed", 0)

    def _bind_callbacks(self) -> None:
        self.view.on_add_files = self.add_files
        self.view.on_remove_selected = self.remove_selected
        self.view.on_clear_all = self.clear_all
        self.view.on_move_up = self.move_up
        self.view.on_move_down = self.move_down
        self.view.on_sort_files = self.sort_files
        self.view.on_convert = self.start_conversion
        self.view.on_convert_separate = self.start_separate_conversion
        self.view.on_drag_reorder = self.drag_reorder
        self.view.on_cancel = self.cancel_conversion
        self.view.on_new_task_tab = self.create_task_tab
        self.view.on_close_task_tab = self.close_task_tab
        self.view.on_switch_task_tab = self.switch_task_tab

    def _find_task(self, task_name: str) -> Optional[Dict[str, Any]]:
        return self.task_manager.find_task(task_name)

    def _active_task(self) -> Optional[Dict[str, Any]]:
        return self.task_manager.active_task()

    def _active_files(self) -> List[str]:
        return self.task_manager.active_files()

    def _resolve_internal_task_name(self, tab_name: str) -> str:
        return self._display_to_internal_task_name.get(tab_name, tab_name)

    def _finalize_successful_task(self, task_name: str) -> None:
        completed_task = self._find_task(task_name)
        if not completed_task:
            return

        completed_task["status"] = "Completed"
        completed_task["progress"] = 100.0

    def create_task_tab(self) -> None:
        self.task_manager.create_task()
        self._refresh_list()

    def close_task_tab(self, tab_name: str) -> None:
        internal_name = self._resolve_internal_task_name(tab_name)
        task = self._find_task(internal_name)
        if not task:
            return

        is_running = bool(task["is_converting"])
        has_files = len(task["files"]) > 0

        if is_running:
            confirmed = self.view.ask_confirm(
                "Cancel Running Task?",
                f"{tab_name} is currently converting.\n\nClose this tab and cancel the running task?",
            )
            if not confirmed:
                return
            task["cancel_requested"] = True
            task["status"] = "Stopping backend process..."
        elif has_files:
            confirmed = self.view.ask_confirm(
                "Clear Task?",
                f"{tab_name} has files in its list.\n\nClose this tab and clear this task?",
            )
            if not confirmed:
                return

        self.task_manager.remove_task(internal_name)
        self._refresh_list()

    def switch_task_tab(self, tab_name: str) -> None:
        internal_tab_name = self._resolve_internal_task_name(tab_name)
        for task_name in self.task_manager.tab_names():
            if task_name == internal_tab_name:
                self.task_manager.active_task_name = internal_tab_name
                self._refresh_list()
                return

    def add_files(self) -> None:
        self.file_actions.add_files()

    def remove_selected(self) -> None:
        self.file_actions.remove_selected()

    def clear_all(self) -> None:
        self.file_actions.clear_all()

    def move_up(self) -> None:
        self.file_actions.move_up()

    def move_down(self) -> None:
        self.file_actions.move_down()

    def drag_reorder(self, from_index: int, to_index: int) -> None:
        self.file_actions.drag_reorder(from_index, to_index)

    def sort_files(self) -> None:
        self.file_actions.sort_files()

    def start_conversion(self) -> None:
        self.conversion_actions.start_conversion(
            task=self._active_task(),
            files=self._active_files(),
            active_task_name=self.task_manager.active_task_name,
            sanitize_name=self.task_manager.sanitize_name,
        )

    def start_separate_conversion(self) -> None:
        self.conversion_actions.start_separate_conversion(
            task=self._active_task(),
            files=self._active_files(),
        )

    def cancel_conversion(self) -> None:
        self.conversion_actions.cancel_conversion(self._active_task())

    def _update_queue_display(self) -> None:
        running = self.task_manager.running_count()
        self.view.schedule(self.view.update_queue_status, running)

    def _refresh_list(self) -> None:
        task = self._active_task()
        files = self._active_files()
        self._display_to_internal_task_name = compute_tab_display_mapping(
            self.task_manager.tasks,
            self.MAX_TAB_LABEL_LENGTH,
        )
        internal_to_display = {v: k for k, v in self._display_to_internal_task_name.items()}

        active_internal = self.task_manager.active_task_name or "Task 1"
        active_name = internal_to_display.get(active_internal, active_internal)

        running_tabs = [
            internal_to_display.get(name, name)
            for name in self.task_manager.running_task_names()
        ]

        self.view.update_task_tabs(list(self._display_to_internal_task_name.keys()), active_name, running_tabs)
        self.view.update_file_list(files)
        if task:
            status_prefix = "Running" if task["is_converting"] else "Idle"
            self.view.update_status(f"[{task['name']}] {status_prefix} - {task['status']}", task["progress"])
            self.view.update_task_actions(bool(task["is_converting"]))
        self._update_queue_display()
