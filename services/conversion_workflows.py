import os
import shutil
import tempfile
import time
from typing import Callable, List, Optional

from pypdf import PdfWriter

from .conversion_types import ConversionCancelledError


class ConversionWorkflows:
    def __init__(
        self,
        convert_single: Callable[[str, str, Optional[Callable[[], bool]]], None],
    ) -> None:
        self._convert_single = convert_single
        self.temp_dir: Optional[str] = None
        self.temp_pdfs: List[str] = []

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
                self._convert_single(ppt_file, pdf_name, is_cancelled)
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
                self._convert_single(ppt_file, output_path, is_cancelled)
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
