import threading
from typing import Any, Callable, Dict, List, Optional

from services.converter_service import ConversionService
from services.conversion_types import ConversionCancelledError
from views import MainView


class ConversionActions:
    def __init__(
        self,
        view: MainView,
        find_task: Callable[[str], Optional[Dict[str, Any]]],
        is_task_cancelled: Callable[[str], bool],
        refresh_list: Callable[[], None],
        finalize_successful_task: Callable[[str], None],
        open_path: Callable[[str], None],
    ) -> None:
        self._view = view
        self._find_task = find_task
        self._is_task_cancelled = is_task_cancelled
        self._refresh_list = refresh_list
        self._finalize_successful_task = finalize_successful_task
        self._open_path = open_path

    def _start_task_conversion(self, task: Dict[str, Any], task_type: str, output_path: str) -> None:
        if task["is_converting"]:
            self._view.show_warning("Task Running", f"{task['name']} is already converting.")
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

    def start_conversion(
        self,
        task: Optional[Dict[str, Any]],
        files: List[str],
        active_task_name: Optional[str],
        sanitize_name: Callable[[str], str],
    ) -> None:
        if not task:
            return

        task_name = active_task_name or "Task"
        if not files:
            self._view.show_warning("No Files", "Add PowerPoint files first.")
            return

        initialfile = f"{sanitize_name(task_name)}.pdf"
        output_path = self._view.ask_save_file(initialfile=initialfile)
        if not output_path:
            return

        self._start_task_conversion(task, "merge", output_path)

    def start_separate_conversion(self, task: Optional[Dict[str, Any]], files: List[str]) -> None:
        if not task:
            return

        if not files:
            self._view.show_warning("No Files", "Please add PPT files first.")
            return

        output_dir = self._view.ask_directory()
        if not output_dir:
            return

        self._start_task_conversion(task, "separate", output_dir)

    def cancel_conversion(self, task: Optional[Dict[str, Any]]) -> None:
        if not task:
            return
        if not task["is_converting"]:
            self._view.show_warning("Nothing to Cancel", "Current tab has no running conversion.")
            return

        task["cancel_requested"] = True
        task["status"] = "Stopping backend process..."
        self._refresh_list()

    def _do_conversion_task(self, task: Dict[str, Any]) -> None:
        try:
            task_name = task.get("task_name", "Task")
            task_state = self._find_task(task_name)
            if task_state is None:
                return

            service = ConversionService()

            def progress_callback(message: str, progress: float) -> None:
                state = self._find_task(task_name)
                if state and not state["cancel_requested"]:
                    state["status"] = message
                    state["progress"] = progress
                    self._view.schedule(self._refresh_list)

            if task_state["cancel_requested"]:
                raise ConversionCancelledError("Conversion cancelled by user")

            if task["type"] == "merge":
                service.convert_and_merge(
                    ppt_files=task["files"],
                    output_path=task["output_path"],
                    progress_callback=progress_callback,
                    delete_temp=self._view.delete_temp_files,
                    is_cancelled=lambda: self._is_task_cancelled(task_name),
                )

                task_state["status"] = "Completed"
                task_state["progress"] = 100.0
                self._view.schedule(self._refresh_list)
                self._view.schedule(
                    self._view.show_info,
                    "Success",
                    f"{task_name}: PDF created successfully!\n{task['output_path']}",
                )

                if self._view.open_after_conversion:
                    self._open_path(task["output_path"])

                self._finalize_successful_task(task_name)
            else:
                created_files = service.convert_separate(
                    ppt_files=task["files"],
                    output_dir=task["output_path"],
                    progress_callback=progress_callback,
                    is_cancelled=lambda: self._is_task_cancelled(task_name),
                )

                task_state["status"] = "Completed"
                task_state["progress"] = 100.0
                self._view.schedule(self._refresh_list)
                self._view.schedule(
                    self._view.show_info,
                    "Done",
                    f"{task_name}: {len(created_files)} PDF files created in:\n{task['output_path']}",
                )

                if self._view.open_after_conversion:
                    self._open_path(task["output_path"])

                self._finalize_successful_task(task_name)

        except ConversionCancelledError:
            task_state = self._find_task(task.get("task_name", ""))
            if task_state:
                task_state["status"] = "Cancelled"
                task_state["progress"] = 0.0
                self._view.schedule(self._refresh_list)
        except Exception as exc:
            task_state = self._find_task(task.get("task_name", ""))
            if task_state:
                task_state["status"] = "Error"
                task_state["progress"] = 0.0
                self._view.schedule(self._refresh_list)
            self._view.schedule(self._view.show_error, "Error", f"Conversion failed:\n{str(exc)}")
        finally:
            task_name = task.get("task_name", "")
            task_state = self._find_task(task_name)
            if task_state:
                task_state["is_converting"] = False
                task_state["cancel_requested"] = False
            self._view.schedule(self._refresh_list)
