import os
import threading
from typing import Any, Dict, List, Optional

from views import MainView
from services import ConversionService
from services.converter_service import ConversionCancelledError
from .system_ops import open_path
from .task_manager import TaskManager

class AppController:    
    MAX_TAB_LABEL_LENGTH = 18

    def __init__(self, view: MainView):
        self.view = view
        self.service = ConversionService()
        self.task_manager = TaskManager()
        self._display_to_internal_task_name: Dict[str, str] = {}

        self._create_task()
        
        self._bind_callbacks()
        self._refresh_list()
        self._check_backend_availability()

    def _open_path(self, path: str) -> None:
        open_path(path)

    def _check_backend_availability(self):
        try:
            available = self.service.get_available_backends()
            if not available:
                self.view.show_warning("No Backend Found", self.service.get_install_message())
                self.view.update_status("No backend found - install required app", 0)
                return

            backend_name = self.service.get_active_backend_name()
            self.view.update_status(f"Ready - Backend: {backend_name}", 0)
        except Exception as e:
            self.view.show_warning("Backend Check Failed", str(e))
            self.view.update_status("Backend check failed", 0)
    
    def _bind_callbacks(self):
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

    def _create_task(self):
        self.task_manager.create_task()

    def _find_task(self, task_name: str) -> Optional[Dict[str, Any]]:
        return self.task_manager.find_task(task_name)

    def _active_task(self) -> Optional[Dict[str, Any]]:
        return self.task_manager.active_task()

    def _active_files(self) -> List[str]:
        return self.task_manager.active_files()

    def _tab_names(self) -> List[str]:
        return self.task_manager.tab_names()

    def _tab_display_label(self, task: Dict[str, Any]) -> str:
        files = task.get("files", [])
        if not files:
            return str(task["name"])

        first_file = os.path.basename(str(files[0]))
        if len(first_file) <= self.MAX_TAB_LABEL_LENGTH:
            return first_file
        return f"{first_file[:self.MAX_TAB_LABEL_LENGTH]}..."

    def _compute_tab_display_mapping(self) -> Dict[str, str]:
        mapping: Dict[str, str] = {}
        label_counts: Dict[str, int] = {}

        for task in self.task_manager.tasks:
            internal_name = str(task["name"])
            base_label = self._tab_display_label(task)
            count = label_counts.get(base_label, 0) + 1
            label_counts[base_label] = count
            display_label = base_label if count == 1 else f"{base_label} ({count})"
            mapping[display_label] = internal_name

        return mapping

    def _resolve_internal_task_name(self, tab_name: str) -> str:
        return self._display_to_internal_task_name.get(tab_name, tab_name)

    def _sanitize_name(self, value: str) -> str:
        return self.task_manager.sanitize_name(value)

    def _is_task_cancelled(self, task_name: str) -> bool:
        return self.task_manager.is_task_cancelled(task_name)

    def _finalize_successful_task(self, task_name: str) -> None:
        completed_task = self._find_task(task_name)
        if not completed_task:
            return

        completed_task["status"] = "Completed"
        completed_task["progress"] = 100.0

    def create_task_tab(self):
        self._create_task()
        self._refresh_list()

    def _remove_task(self, task_name: str):
        self.task_manager.remove_task(task_name)

    def close_task_tab(self, tab_name: str):
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

        self._remove_task(internal_name)
        self._refresh_list()

    def switch_task_tab(self, tab_name: str):
        internal_tab_name = self._resolve_internal_task_name(tab_name)
        for task_name in self._tab_names():
            if task_name == internal_tab_name:
                self.task_manager.active_task_name = internal_tab_name
                self._refresh_list()
                return
    
    def add_files(self):
        active_task = self._active_task()
        if active_task and active_task["is_converting"]:
            self.view.show_warning("Task Running", "Cannot edit files while this task is converting.")
            return

        new_files = self.view.ask_open_files()
        files = self._active_files()
        
        for f in new_files:
            if f not in files:
                files.append(f)
        
        self._refresh_list()
    
    def remove_selected(self):
        active_task = self._active_task()
        if active_task and active_task["is_converting"]:
            self.view.show_warning("Task Running", "Cannot edit files while this task is converting.")
            return
        i = self.view.get_selected_index()
        files = self._active_files()
        if i is not None:
            files.pop(i)
            self._refresh_list()
    
    def clear_all(self):
        active_task = self._active_task()
        if active_task and active_task["is_converting"]:
            self.view.show_warning("Task Running", "Cannot edit files while this task is converting.")
            return

        if not active_task:
            return

        files = active_task["files"]
        files.clear()
        active_task["status"] = "Ready - Add files to get started"
        active_task["progress"] = 0.0
        self._refresh_list()
    
    def move_up(self):
        active_task = self._active_task()
        if active_task and active_task["is_converting"]:
            self.view.show_warning("Task Running", "Cannot reorder files while this task is converting.")
            return

        i = self.view.get_selected_index()
        files = self._active_files()
        if i is not None and i > 0:
            files[i], files[i-1] = files[i-1], files[i]
            self._refresh_list()
            self.view.set_selection(i - 1)
    
    def move_down(self):
        active_task = self._active_task()
        if active_task and active_task["is_converting"]:
            self.view.show_warning("Task Running", "Cannot reorder files while this task is converting.")
            return

        i = self.view.get_selected_index()
        files = self._active_files()
        if i is not None and i < len(files) - 1:
            files[i], files[i+1] = files[i+1], files[i]
            self._refresh_list()
            self.view.set_selection(i + 1)
    
    def drag_reorder(self, from_index: int, to_index: int):
        active_task = self._active_task()
        if active_task and active_task["is_converting"]:
            return

        files = self._active_files()
        if from_index is not None and 0 <= from_index < len(files):
            i = files.pop(from_index)
            files.insert(to_index, i)
            self._refresh_list()
            self.view.set_selection(to_index)

    def sort_files(self):
        active_task = self._active_task()
        if active_task and active_task["is_converting"]:
            self.view.show_warning("Task Running", "Cannot sort files while this task is converting.")
            return
        files = self._active_files()
        files.sort(key=lambda path: os.path.basename(path).lower())
        self._refresh_list()

    def _start_task_conversion(self, task: Dict[str, Any], task_type: str, output_path: str):
        if task["is_converting"]:
            self.view.show_warning("Task Running", f"{task['name']} is already converting.")
            return

        task["is_converting"] = True
        task["cancel_requested"] = False
        task["status"] = "Starting conversion..."
        task["progress"] = 0.0
        self._refresh_list()

        task_payload = {
            "task_name": task["name"],
            "type": task_type,
            "files": list(task["files"]),
            "output_path": output_path,
        }
        thread = threading.Thread(target=self._do_conversion_task, args=(task_payload,), daemon=True)
        thread.start()
    
    def start_conversion(self):
        task = self._active_task()
        if not task:
            return

        files = self._active_files()
        task_name = self.task_manager.active_task_name or "Task"
        if not files:
            self.view.show_warning("No Files", "Add PowerPoint files first.")
            return

        initialfile = f"{self._sanitize_name(task_name)}.pdf"
        dPath = self.view.ask_save_file(initialfile=initialfile)
        if not dPath:
            return

        self._start_task_conversion(task, "merge", dPath)
    
    def _do_conversion_task(self, task: dict):
        """Execute a single conversion task."""
        try:
            task_name = task.get("task_name", "Task")
            task_state = self._find_task(task_name)
            if task_state is None:
                return
            service = ConversionService()

            def progress_callback(message: str, progress: float):
                state = self._find_task(task_name)
                if state and not state["cancel_requested"]:
                    state["status"] = message
                    state["progress"] = progress
                    self.view.schedule(self._refresh_list)

            if task_state["cancel_requested"]:
                raise ConversionCancelledError("Conversion cancelled by user")
            
            if task["type"] == "merge":
                service.convert_and_merge(
                    ppt_files=task["files"],
                    output_path=task["output_path"],
                    progress_callback=progress_callback,
                    delete_temp=self.view.delete_temp_files,
                    is_cancelled=lambda: self._is_task_cancelled(task_name),
                )

                task_state["status"] = "Completed"
                task_state["progress"] = 100.0
                self.view.schedule(self._refresh_list)
                
                self.view.schedule(
                    self.view.show_info,
                    "Success",
                    f"{task_name}: PDF created successfully!\n{task['output_path']}"
                )
                
                # Open the PDF if requested
                if self.view.open_after_conversion:
                    self._open_path(task["output_path"])

                self._finalize_successful_task(task_name)
            
            else:  # separate
                created_files = service.convert_separate(
                    ppt_files=task["files"],
                    output_dir=task["output_path"],
                    progress_callback=progress_callback,
                    is_cancelled=lambda: self._is_task_cancelled(task_name),
                )

                task_state["status"] = "Completed"
                task_state["progress"] = 100.0
                self.view.schedule(self._refresh_list)
                
                self.view.schedule(
                    self.view.show_info,
                    "Done",
                    f"{task_name}: {len(created_files)} PDF files created in:\n{task['output_path']}"
                )
                
                # Open the folder if requested
                if self.view.open_after_conversion:
                    self._open_path(task["output_path"])

                self._finalize_successful_task(task_name)
        
        except ConversionCancelledError:
            task_state = self._find_task(task.get("task_name", ""))
            if task_state:
                task_state["status"] = "Cancelled"
                task_state["progress"] = 0.0
                self.view.schedule(self._refresh_list)
        except Exception as e:
            task_state = self._find_task(task.get("task_name", ""))
            if task_state:
                task_state["status"] = "Error"
                task_state["progress"] = 0.0
                self.view.schedule(self._refresh_list)
            self.view.schedule(
                self.view.show_error,
                "Error",
                f"Conversion failed:\n{str(e)}"
            )
        
        finally:
            task_name = task.get("task_name", "")
            task_state = self._find_task(task_name)
            if task_state:
                task_state["is_converting"] = False
                task_state["cancel_requested"] = False
            self.view.schedule(self._refresh_list)
    
    def start_separate_conversion(self):
        task = self._active_task()
        if not task:
            return

        files = self._active_files()
        if not files:
            self.view.show_warning("No Files", "Please add PPT files first.")
            return
        
        output_dir = self.view.ask_directory()
        if not output_dir:
            return
        self._start_task_conversion(task, "separate", output_dir)
    
    def cancel_conversion(self):
        task = self._active_task()
        if not task:
            return
        if not task["is_converting"]:
            self.view.show_warning("Nothing to Cancel", "Current tab has no running conversion.")
            return

        task["cancel_requested"] = True
        task["status"] = "Stopping backend process..."
        self._refresh_list()
    
    def _update_queue_display(self):
        running = self.task_manager.running_count()
        self.view.schedule(self.view.update_queue_status, running)
    
    def _refresh_list(self):
        task = self._active_task()
        files = self._active_files()
        self._display_to_internal_task_name = self._compute_tab_display_mapping()
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
