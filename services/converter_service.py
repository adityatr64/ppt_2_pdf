from typing import Callable, List, Optional

from .backend_converters import BackendConverters
from .backend_support import BackendSupport
from .conversion_types import ConversionBackend, ConversionCancelledError
from .conversion_workflows import ConversionWorkflows


class ConversionService:
    def __init__(self):
        self._backend_support = BackendSupport()
        self._backend_converters = BackendConverters(self._backend_support)
        self._workflows = ConversionWorkflows(self.ppt_to_pdf)

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
        self._backend_converters.convert(input_path, output_path, is_cancelled)

    def merge_pdfs(
        self,
        pdf_list: List[str],
        output_path: str,
        is_cancelled: Optional[Callable[[], bool]] = None,
    ) -> None:
        self._workflows.merge_pdfs(pdf_list, output_path, is_cancelled)

    def convert_and_merge(
        self,
        ppt_files: List[str],
        output_path: str,
        progress_callback: Optional[Callable[[str, float], None]] = None,
        delete_temp: bool = True,
        is_cancelled: Optional[Callable[[], bool]] = None,
    ) -> None:
        self._workflows.convert_and_merge(
            ppt_files=ppt_files,
            output_path=output_path,
            progress_callback=progress_callback,
            delete_temp=delete_temp,
            is_cancelled=is_cancelled,
        )

    def convert_separate(
        self,
        ppt_files: List[str],
        output_dir: str,
        progress_callback: Optional[Callable[[str, float], None]] = None,
        is_cancelled: Optional[Callable[[], bool]] = None,
    ) -> List[str]:
        return self._workflows.convert_separate(
            ppt_files=ppt_files,
            output_dir=output_dir,
            progress_callback=progress_callback,
            is_cancelled=is_cancelled,
        )

    def cleanup_temp_files(self) -> None:
        self._workflows.cleanup_temp_files()


__all__ = [
    "ConversionService",
    "ConversionBackend",
    "ConversionCancelledError",
]
