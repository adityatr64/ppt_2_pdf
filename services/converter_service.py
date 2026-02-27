import os
import shutil
import sys
import tempfile
import time
from typing import Callable, List, Optional

from pypdf import PdfWriter
# Shit's been documented using AI can have error in docs itself
# Nvm That was a older version, only windows version written truly by human
# Linux and Macos are prone to errors cos i don't trust AI to do a decent job

from .backend_support import BackendSupport
from .conversion_runtime import (
    powerpoint_worker,
    run_cancellable_subprocess,
    run_cancellable_worker,
    wps_worker,
)
from .conversion_types import ConversionBackend, ConversionCancelledError


class ConversionService:
    def __init__(self):
        self.temp_dir: Optional[str] = None
        self.temp_pdfs: List[str] = []
        self._backend_support = BackendSupport()

    def get_available_backends(self) -> List[ConversionBackend]:
        return self._backend_support.get_available_backends()

    def get_install_message(self) -> str:
        return self._backend_support.get_install_message()

    def get_active_backend_name(self) -> str:
        return self._backend_support.get_active_backend_name()

    def ppt_to_pdf(
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

    def merge_pdfs(
        self,
        pdf_list: List[str],
        output_path: str,
        is_cancelled: Optional[Callable[[], bool]] = None,
    ) -> None:
        merger = PdfWriter()
        try:
            for pdf in pdf_list:
                if is_cancelled and is_cancelled():
                    raise ConversionCancelledError("Conversion cancelled")
                merger.append(pdf)

            if is_cancelled and is_cancelled():
                raise ConversionCancelledError("Conversion cancelled")

            merger.write(output_path)
        finally:
            merger.close()

    def convert_and_merge(
        self,
        ppt_files: List[str],
        output_path: str,
        progress_callback: Optional[Callable[[str, float], None]] = None,
        delete_temp: bool = True,
        is_cancelled: Optional[Callable[[], bool]] = None,
    ) -> None:
        self.temp_pdfs = []
        self.temp_dir = tempfile.mkdtemp()

        try:
            total_files = len(ppt_files)

            for i, ppt_file in enumerate(ppt_files):
                if is_cancelled and is_cancelled():
                    raise ConversionCancelledError("Conversion cancelled")

                if progress_callback:
                    progress_callback(
                        f"Converting {i + 1}/{total_files}: {os.path.basename(ppt_file)}",
                        (i / total_files) * 80,
                    )

                pdf_name = os.path.join(self.temp_dir, f"temp_{i}.pdf")
                self.ppt_to_pdf(ppt_file, pdf_name, is_cancelled=is_cancelled)
                self.temp_pdfs.append(pdf_name)

                if is_cancelled and is_cancelled():
                    raise ConversionCancelledError("Conversion cancelled")

            if progress_callback:
                progress_callback("Merging PDFs...", 90)

            self.merge_pdfs(self.temp_pdfs, output_path, is_cancelled=is_cancelled)

            if progress_callback:
                progress_callback("Conversion complete!", 100)
        except ConversionCancelledError:
            if os.path.exists(output_path):
                try:
                    os.remove(output_path)
                except OSError:
                    pass
            raise
        finally:
            if delete_temp:
                self.cleanup_temp_files()

    def convert_separate(
        self,
        ppt_files: List[str],
        output_dir: str,
        progress_callback: Optional[Callable[[str, float], None]] = None,
        is_cancelled: Optional[Callable[[], bool]] = None,
    ) -> List[str]:
        created_pdfs = []
        total_files = len(ppt_files)
        expected_outputs: List[str] = []

        for ppt_file in ppt_files:
            base_name = os.path.splitext(os.path.basename(ppt_file))[0]
            output_name = f"{base_name}_converted_pdf.pdf"
            expected_outputs.append(os.path.join(output_dir, output_name))

        try:
            for i, ppt_file in enumerate(ppt_files):
                if is_cancelled and is_cancelled():
                    raise ConversionCancelledError("Conversion cancelled")

                if progress_callback:
                    progress_callback(
                        f"Converting {i + 1}/{total_files}: {os.path.basename(ppt_file)}",
                        (i / total_files) * 100,
                    )

                output_path = expected_outputs[i]

                self.ppt_to_pdf(ppt_file, output_path, is_cancelled=is_cancelled)
                created_pdfs.append(output_path)

                if is_cancelled and is_cancelled():
                    raise ConversionCancelledError("Conversion cancelled")

            if progress_callback:
                progress_callback("All files converted!", 100)

            return created_pdfs
        except ConversionCancelledError:
            time.sleep(1.0)
            for expected_output in expected_outputs:
                if os.path.exists(expected_output):
                    try:
                        os.remove(expected_output)
                    except OSError:
                        pass
            raise

    def cleanup_temp_files(self) -> None:
        for pdf in self.temp_pdfs:
            try:
                os.remove(pdf)
            except OSError:
                pass

        if self.temp_dir:
            shutil.rmtree(self.temp_dir, ignore_errors=True)

        self.temp_pdfs = []
        self.temp_dir = None
