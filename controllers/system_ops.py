import os
import subprocess
import sys


def open_path(path: str) -> None:
    if sys.platform == "win32":
        os.startfile(path)
        return
    if sys.platform == "darwin":
        subprocess.run(["open", path], check=False)
        return
    subprocess.run(["xdg-open", path], check=False)
