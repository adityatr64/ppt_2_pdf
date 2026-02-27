import multiprocessing as mp
import os
import subprocess
import time
from typing import Callable, List, Optional

from .conversion_types import ConversionCancelledError


def powerpoint_worker(input_path: str, output_path: str, result_queue: "mp.Queue") -> None:
    try:
        import comtypes.client

        powerpoint = None
        deck = None
        try:
            powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
            deck = powerpoint.Presentations.Open(os.path.abspath(input_path), WithWindow=False)
            deck.SaveAs(os.path.abspath(output_path), 32)
        finally:
            if deck:
                deck.Close()
            if powerpoint:
                powerpoint.Quit()

        result_queue.put(None)
    except Exception as exc:
        result_queue.put(str(exc))


def wps_worker(input_path: str, output_path: str, result_queue: "mp.Queue") -> None:
    try:
        import comtypes.client

        app = None
        deck = None
        try:
            app = comtypes.client.CreateObject("KWPP.Application")
            deck = app.Presentations.Open(os.path.abspath(input_path), False, False, False)
            if hasattr(deck, "ExportAsFixedFormat"):
                deck.ExportAsFixedFormat(os.path.abspath(output_path), 2)
            else:
                deck.SaveAs(os.path.abspath(output_path), 32)
        finally:
            if deck:
                deck.Close()
            if app:
                app.Quit()

        result_queue.put(None)
    except Exception as exc:
        result_queue.put(str(exc))


def run_cancellable_subprocess(
    cmd: List[str],
    timeout: int,
    is_cancelled: Optional[Callable[[], bool]] = None,
) -> subprocess.CompletedProcess:
    process = subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )

    start_time = time.time()
    while True:
        if is_cancelled and is_cancelled():
            process.terminate()
            try:
                process.wait(timeout=5)
            except subprocess.TimeoutExpired:
                process.kill()
            raise ConversionCancelledError("Conversion cancelled")

        return_code = process.poll()
        if return_code is not None:
            stdout, stderr = process.communicate()
            return subprocess.CompletedProcess(cmd, return_code, stdout, stderr)

        if time.time() - start_time > timeout:
            process.terminate()
            try:
                process.wait(timeout=5)
            except subprocess.TimeoutExpired:
                process.kill()
            raise RuntimeError("Conversion timed out")

        time.sleep(0.2)


def run_cancellable_worker(
    worker_func: Callable[..., None],
    input_path: str,
    output_path: str,
    is_cancelled: Optional[Callable[[], bool]] = None,
) -> None:
    ctx = mp.get_context("spawn")
    result_queue: mp.Queue = ctx.Queue()
    process = ctx.Process(target=worker_func, args=(input_path, output_path, result_queue), daemon=True)
    process.start()

    try:
        while process.is_alive():
            if is_cancelled and is_cancelled():
                process.terminate()
                process.join(timeout=5)
                raise ConversionCancelledError("Conversion cancelled")
            time.sleep(0.2)

        process.join(timeout=5)
        if process.exitcode not in (0, None):
            error_message = None
            if not result_queue.empty():
                error_message = result_queue.get()
            raise RuntimeError(error_message or "Backend conversion process failed")

        if not result_queue.empty():
            error_message = result_queue.get()
            if error_message:
                raise RuntimeError(error_message)
    finally:
        if process.is_alive():
            process.terminate()
            process.join(timeout=1)
