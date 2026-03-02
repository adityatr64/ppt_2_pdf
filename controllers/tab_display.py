import os
from typing import Any, Dict, List


def tab_display_label(task: Dict[str, Any], max_label_length: int) -> str:
    files = task.get("files", [])
    if not files:
        return str(task["name"])

    first_file = os.path.basename(str(files[0]))
    if len(first_file) <= max_label_length:
        return first_file
    return f"{first_file[:max_label_length]}..."


def compute_tab_display_mapping(tasks: List[Dict[str, Any]], max_label_length: int) -> Dict[str, str]:
    mapping: Dict[str, str] = {}
    label_counts: Dict[str, int] = {}

    for task in tasks:
        internal_name = str(task["name"])
        base_label = tab_display_label(task, max_label_length)
        count = label_counts.get(base_label, 0) + 1
        label_counts[base_label] = count
        display_label = base_label if count == 1 else f"{base_label} ({count})"
        mapping[display_label] = internal_name

    return mapping
