from typing import Any, Dict, List, Optional


class TaskManager:
    def __init__(self) -> None:
        self._tasks: List[Dict[str, Any]] = []
        self._active_task_name: Optional[str] = None
        self._task_counter = 0

    @property
    def tasks(self) -> List[Dict[str, Any]]:
        return self._tasks

    @property
    def active_task_name(self) -> Optional[str]:
        return self._active_task_name

    @active_task_name.setter
    def active_task_name(self, value: Optional[str]) -> None:
        self._active_task_name = value

    def create_task(self) -> Dict[str, Any]:
        self._task_counter += 1
        task_name = f"Task {self._task_counter}"
        task = {
            "name": task_name,
            "files": [],
            "is_converting": False,
            "cancel_requested": False,
            "status": "Ready - Add files to get started",
            "progress": 0.0,
        }
        self._tasks.append(task)
        self._active_task_name = task_name
        return task

    def find_task(self, task_name: str) -> Optional[Dict[str, Any]]:
        for task in self._tasks:
            if task["name"] == task_name:
                return task
        return None

    def active_task(self) -> Optional[Dict[str, Any]]:
        if self._active_task_name is None:
            return None
        return self.find_task(self._active_task_name)

    def active_files(self) -> List[str]:
        task = self.active_task()
        if not task:
            return []
        return task["files"]

    def tab_names(self) -> List[str]:
        return [str(task["name"]) for task in self._tasks]

    def sanitize_name(self, value: str) -> str:
        clean = "".join(ch if ch.isalnum() or ch in ("-", "_") else "_" for ch in value.strip())
        return clean.strip("_") or "task"

    def is_task_cancelled(self, task_name: str) -> bool:
        task = self.find_task(task_name)
        if not task:
            return True
        return bool(task["cancel_requested"])

    def remove_task(self, task_name: str) -> None:
        self._tasks = [task for task in self._tasks if task["name"] != task_name]
        if not self._tasks:
            self.create_task()
            return
        if self._active_task_name == task_name:
            self._active_task_name = str(self._tasks[0]["name"])

    def running_task_names(self) -> List[str]:
        return [str(task["name"]) for task in self._tasks if task["is_converting"]]

    def running_count(self) -> int:
        return sum(1 for task in self._tasks if task["is_converting"])
