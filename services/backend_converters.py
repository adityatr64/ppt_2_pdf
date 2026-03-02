import os
import shutil
import sys
import tempfile
from typing import Callable, Optional

from .backend_support import BackendSupport
from .conversion_runtime import (
    powerpoint_worker,
    run_cancellable_subprocess,
    run_cancellable_worker,
    wps_worker,
)
from .conversion_types import ConversionBackend, ConversionCancelledError


class BackendConverters:
    def __init__(self, backend_support: Optional[BackendSupport] = None) -> None:
        self._backend_support = backend_support or BackendSupport()

    def convert(
        self,
        input_path: str,
        output_path: str,
        is_cancelled: Optional[Callable[[], bool]] = None,
    ) -> None:
        backend = self._backend_support.get_active_backend()

        if backend == ConversionBackend.POWERPOINT:
            self._convert_with_powerpoint(input_path, output_path, is_cancelled=is_cancelled)
            return
        if backend == ConversionBackend.WPS:
            self._convert_with_wps(input_path, output_path, is_cancelled=is_cancelled)
            return
        if backend == ConversionBackend.LIBREOFFICE:
            self._convert_with_libreoffice(input_path, output_path, is_cancelled=is_cancelled)
            return
        if backend == ConversionBackend.ONLYOFFICE:
            self._convert_with_onlyoffice(input_path, output_path, is_cancelled=is_cancelled)
            return
        if backend == ConversionBackend.KEYNOTE:
            self._convert_with_keynote(input_path, output_path, is_cancelled=is_cancelled)
            return

        raise RuntimeError(f"Unsupported backend: {backend.value}")

    def _convert_with_powerpoint(
        self,
        input_path: str,
        output_path: str,
        is_cancelled: Optional[Callable[[], bool]] = None,
    ) -> None:
        if sys.platform != "win32":
            raise RuntimeError("PowerPoint backend is only supported on Windows")
        run_cancellable_worker(powerpoint_worker, input_path, output_path, is_cancelled=is_cancelled)

    def _convert_with_wps(
        self,
        input_path: str,
        output_path: str,
        is_cancelled: Optional[Callable[[], bool]] = None,
    ) -> None:
        if sys.platform != "win32":
            raise RuntimeError("WPS backend is only supported on Windows")

        try:
            run_cancellable_worker(wps_worker, input_path, output_path, is_cancelled=is_cancelled)
            return
        except ConversionCancelledError:
            raise
        except Exception:
            pass

        raise RuntimeError(
            "WPS Office detected but automated PDF export failed. "
            "Please install PowerPoint or LibreOffice for reliable automation."
        )

    def _convert_with_libreoffice(
        self,
        input_path: str,
        output_path: str,
        is_cancelled: Optional[Callable[[], bool]] = None,
    ) -> None:
        soffice = self._backend_support.find_libreoffice()
        if not soffice:
            raise RuntimeError("LibreOffice executable not found")

        input_abs = os.path.abspath(input_path)
        with tempfile.TemporaryDirectory() as temp_out_dir:
            cmd = [
                soffice,
                "--headless",
                "--invisible",
                "--convert-to",
                "pdf",
                "--outdir",
                temp_out_dir,
                input_abs,
            ]

            result = run_cancellable_subprocess(cmd, timeout=180, is_cancelled=is_cancelled)
            if result.returncode != 0:
                raise RuntimeError(f"LibreOffice conversion failed: {result.stderr.strip() or result.stdout.strip()}")

            base_name = os.path.splitext(os.path.basename(input_path))[0]
            generated_pdf = os.path.join(temp_out_dir, f"{base_name}.pdf")
            if not os.path.exists(generated_pdf):
                raise RuntimeError("LibreOffice conversion finished but no PDF was generated")

            shutil.move(generated_pdf, os.path.abspath(output_path))

    def _convert_with_onlyoffice(
        self,
        input_path: str,
        output_path: str,
        is_cancelled: Optional[Callable[[], bool]] = None,
    ) -> None:
        exe = self._backend_support.find_onlyoffice()
        if not exe:
            raise RuntimeError("ONLYOFFICE executable not found")

        with tempfile.TemporaryDirectory() as temp_out_dir:
            cmd = [
                exe,
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                temp_out_dir,
                os.path.abspath(input_path),
            ]

            result = run_cancellable_subprocess(cmd, timeout=180, is_cancelled=is_cancelled)
            if result.returncode != 0:
                raise RuntimeError(
                    "ONLYOFFICE conversion failed. Desktop Editors CLI may not support conversion on this build. "
                    f"Details: {result.stderr.strip() or result.stdout.strip()}"
                )

            base_name = os.path.splitext(os.path.basename(input_path))[0]
            generated_pdf = os.path.join(temp_out_dir, f"{base_name}.pdf")
            if not os.path.exists(generated_pdf):
                raise RuntimeError("ONLYOFFICE conversion finished but no PDF was generated")

            shutil.move(generated_pdf, os.path.abspath(output_path))

    def _convert_with_keynote(
        self,
        input_path: str,
        output_path: str,
        is_cancelled: Optional[Callable[[], bool]] = None,
    ) -> None:
        if sys.platform != "darwin":
            raise RuntimeError("Keynote backend is only supported on macOS")

        script = [
            "tell application \"Keynote\"",
            "activate",
            f"set inFile to POSIX file \"{os.path.abspath(input_path)}\"",
            "set theDoc to open inFile",
            f"export theDoc to POSIX file \"{os.path.abspath(output_path)}\" as PDF",
            "close theDoc saving no",
            "end tell",
        ]

        result = run_cancellable_subprocess(
            ["osascript", "-e", "\n".join(script)],
            timeout=180,
            is_cancelled=is_cancelled,
        )
        if result.returncode != 0:
            raise RuntimeError(f"Keynote conversion failed: {result.stderr.strip() or result.stdout.strip()}")
