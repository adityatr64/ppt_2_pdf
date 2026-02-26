import os
import sys
import shutil
import tempfile
import subprocess
from enum import Enum
from pypdf import PdfWriter
from typing import List, Callable, Optional

# Shit's been documented using AI can have error in docs itself
# Nvm That was a older version newwer version code is only written truly by human
# Linux and Macos are prone to errors
class ConversionBackend(Enum):
    POWERPOINT = "powerpoint"
    WPS = "wps"
    LIBREOFFICE = "libreoffice"
    ONLYOFFICE = "onlyoffice"
    KEYNOTE = "keynote"


class ConversionService:
    LIBREOFFICE_PATHS = {
        "win32": [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ],
        "darwin": [
            "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        ],
        "linux": [
            "/usr/bin/libreoffice",
            "/usr/bin/soffice",
            "/snap/bin/libreoffice",
        ],
    }

    WPS_PATHS = [
        r"C:\Program Files\WPS Office\12.2.0.17119\office6\wpp.exe",
        r"C:\Program Files\WPS Office\office6\wpp.exe",
        r"C:\Program Files (x86)\WPS Office\office6\wpp.exe",
    ]

    ONLYOFFICE_PATHS = {
        "win32": [
            r"C:\Program Files\ONLYOFFICE\DesktopEditors\DesktopEditors.exe",
            r"C:\Program Files (x86)\ONLYOFFICE\DesktopEditors\DesktopEditors.exe",
        ],
        "darwin": [
            "/Applications/ONLYOFFICE.app/Contents/MacOS/DesktopEditors",
        ],
        "linux": [
            "/usr/bin/onlyoffice-desktopeditors",
            "/usr/bin/desktopeditors",
        ],
    }

    def __init__(self):
        self.temp_dir: Optional[str] = None
        self.temp_pdfs: List[str] = []
        self._detected_backend: Optional[ConversionBackend] = None
        self._libreoffice_path: Optional[str] = None
        self._onlyoffice_path: Optional[str] = None

    def _platform_key(self) -> str:
        if sys.platform == "win32":
            return "win32"
        if sys.platform == "darwin":
            return "darwin"
        return "linux"

    def _find_executable(self, names: List[str], extra_paths: Optional[List[str]] = None) -> Optional[str]:
        for name in names:
            path = shutil.which(name)
            if path:
                return path

        if extra_paths:
            for path in extra_paths:
                if os.path.isfile(path):
                    return path

        return None

    def _check_powerpoint_available(self) -> bool:
        if sys.platform != "win32":
            return False

        try:
            import comtypes.client

            app = comtypes.client.CreateObject("Powerpoint.Application")
            app.Quit()
            return True
        except Exception:
            return False

    def _check_wps_available(self) -> bool:
        if sys.platform != "win32":
            return False

        try:
            import comtypes.client

            app = comtypes.client.CreateObject("KWPP.Application")
            app.Quit()
            return True
        except Exception:
            pass

        return self._find_executable(["wpp", "wps"], self.WPS_PATHS) is not None

    def _find_libreoffice(self) -> Optional[str]:
        if self._libreoffice_path:
            return self._libreoffice_path

        platform_paths = self.LIBREOFFICE_PATHS.get(self._platform_key(), [])
        self._libreoffice_path = self._find_executable(["soffice", "libreoffice"], platform_paths)
        return self._libreoffice_path

    def _find_onlyoffice(self) -> Optional[str]:
        if self._onlyoffice_path:
            return self._onlyoffice_path

        platform_paths = self.ONLYOFFICE_PATHS.get(self._platform_key(), [])
        self._onlyoffice_path = self._find_executable(
            ["onlyoffice-desktopeditors", "desktopeditors", "DesktopEditors"],
            platform_paths,
        )
        return self._onlyoffice_path

    def _check_keynote_available(self) -> bool:
        if sys.platform != "darwin":
            return False
        return os.path.isdir("/Applications/Keynote.app")

    def get_available_backends(self) -> List[ConversionBackend]:
        if sys.platform == "win32":
            ordered = [
                ConversionBackend.POWERPOINT,
                ConversionBackend.WPS,
                ConversionBackend.LIBREOFFICE,
                ConversionBackend.ONLYOFFICE,
            ]
        elif sys.platform == "darwin":
            ordered = [
                ConversionBackend.KEYNOTE,
                ConversionBackend.LIBREOFFICE,
                ConversionBackend.ONLYOFFICE,
            ]
        else:
            ordered = [
                ConversionBackend.LIBREOFFICE,
                ConversionBackend.ONLYOFFICE,
            ]

        available: List[ConversionBackend] = []
        for backend in ordered:
            if backend == ConversionBackend.POWERPOINT and self._check_powerpoint_available():
                available.append(backend)
            elif backend == ConversionBackend.WPS and self._check_wps_available():
                available.append(backend)
            elif backend == ConversionBackend.LIBREOFFICE and self._find_libreoffice():
                available.append(backend)
            elif backend == ConversionBackend.ONLYOFFICE and self._find_onlyoffice():
                available.append(backend)
            elif backend == ConversionBackend.KEYNOTE and self._check_keynote_available():
                available.append(backend)

        return available

    def get_install_message(self) -> str:
        if sys.platform == "win32":
            choices = "Microsoft PowerPoint, WPS Office, LibreOffice, or ONLYOFFICE"
        elif sys.platform == "darwin":
            choices = "Keynote, LibreOffice, or ONLYOFFICE"
        else:
            choices = "LibreOffice or ONLYOFFICE"

        return (
            "No supported converter backend found.\n\n"
            f"Please install one of: {choices}."
        )

    def _get_active_backend(self) -> ConversionBackend:
        if self._detected_backend:
            return self._detected_backend

        available = self.get_available_backends()
        if not available:
            raise RuntimeError(self.get_install_message())

        self._detected_backend = available[0]
        return self._detected_backend

    def get_active_backend_name(self) -> str:
        backend = self._get_active_backend()
        if backend == ConversionBackend.POWERPOINT:
            return "Microsoft PowerPoint"
        if backend == ConversionBackend.WPS:
            return "WPS Office"
        if backend == ConversionBackend.LIBREOFFICE:
            return "LibreOffice"
        if backend == ConversionBackend.ONLYOFFICE:
            return "ONLYOFFICE"
        if backend == ConversionBackend.KEYNOTE:
            return "Apple Keynote"
        return backend.value

    def ppt_to_pdf(self, input_path: str, output_path: str) -> None:
        backend = self._get_active_backend()

        if backend == ConversionBackend.POWERPOINT:
            self._convert_with_powerpoint(input_path, output_path)
            return
        if backend == ConversionBackend.WPS:
            self._convert_with_wps(input_path, output_path)
            return
        if backend == ConversionBackend.LIBREOFFICE:
            self._convert_with_libreoffice(input_path, output_path)
            return
        if backend == ConversionBackend.ONLYOFFICE:
            self._convert_with_onlyoffice(input_path, output_path)
            return
        if backend == ConversionBackend.KEYNOTE:
            self._convert_with_keynote(input_path, output_path)
            return

        raise RuntimeError(f"Unsupported backend: {backend.value}")

    def _convert_with_powerpoint(self, input_path: str, output_path: str) -> None:
        import comtypes.client

        powerpoint = None
        deck = None
        try:
            powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
            deck = powerpoint.Presentations.Open(os.path.abspath(input_path), WithWindow=False)
            deck.SaveAs(os.path.abspath(output_path), 32)  # 32 = ppSaveAsPDF
        finally:
            if deck:
                deck.Close()
            if powerpoint:
                powerpoint.Quit()

    def _convert_with_wps(self, input_path: str, output_path: str) -> None:
        if sys.platform != "win32":
            raise RuntimeError("WPS backend is only supported on Windows")

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
                return
            finally:
                if deck:
                    deck.Close()
                if app:
                    app.Quit()
        except Exception:
            pass

        raise RuntimeError(
            "WPS Office detected but automated PDF export failed. "
            "Please install PowerPoint or LibreOffice for reliable automation."
        )

    def _convert_with_libreoffice(self, input_path: str, output_path: str) -> None:
        soffice = self._find_libreoffice()
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

            result = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
            if result.returncode != 0:
                raise RuntimeError(f"LibreOffice conversion failed: {result.stderr.strip() or result.stdout.strip()}")

            base_name = os.path.splitext(os.path.basename(input_path))[0]
            generated_pdf = os.path.join(temp_out_dir, f"{base_name}.pdf")
            if not os.path.exists(generated_pdf):
                raise RuntimeError("LibreOffice conversion finished but no PDF was generated")

            shutil.move(generated_pdf, os.path.abspath(output_path))

    def _convert_with_onlyoffice(self, input_path: str, output_path: str) -> None:
        exe = self._find_onlyoffice()
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

            result = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
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

    def _convert_with_keynote(self, input_path: str, output_path: str) -> None:
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

        result = subprocess.run(["osascript", "-e", "\n".join(script)], capture_output=True, text=True, timeout=180)
        if result.returncode != 0:
            raise RuntimeError(f"Keynote conversion failed: {result.stderr.strip() or result.stdout.strip()}")

    def merge_pdfs(self, pdf_list: List[str], output_path: str) -> None:
        merger = PdfWriter()
        for pdf in pdf_list:
            merger.append(pdf)
        merger.write(output_path)
        merger.close()

    def convert_and_merge(
        self,
        ppt_files: List[str],
        output_path: str,
        progress_callback: Optional[Callable[[str, float], None]] = None,
        delete_temp: bool = True,
    ) -> None:
        self.temp_pdfs = []
        self.temp_dir = tempfile.mkdtemp()

        try:
            total_files = len(ppt_files)

            for i, ppt_file in enumerate(ppt_files):
                if progress_callback:
                    progress_callback(
                        f"Converting {i + 1}/{total_files}: {os.path.basename(ppt_file)}",
                        (i / total_files) * 80,
                    )

                pdf_name = os.path.join(self.temp_dir, f"temp_{i}.pdf")
                self.ppt_to_pdf(ppt_file, pdf_name)
                self.temp_pdfs.append(pdf_name)

            if progress_callback:
                progress_callback("Merging PDFs...", 90)

            self.merge_pdfs(self.temp_pdfs, output_path)

            if progress_callback:
                progress_callback("Conversion complete!", 100)
        finally:
            if delete_temp:
                self.cleanup_temp_files()

    def convert_separate(
        self,
        ppt_files: List[str],
        output_dir: str,
        progress_callback: Optional[Callable[[str, float], None]] = None,
    ) -> List[str]:
        created_pdfs = []
        total_files = len(ppt_files)

        for i, ppt_file in enumerate(ppt_files):
            if progress_callback:
                progress_callback(
                    f"Converting {i + 1}/{total_files}: {os.path.basename(ppt_file)}",
                    (i / total_files) * 100,
                )

            base_name = os.path.splitext(os.path.basename(ppt_file))[0]
            output_name = f"{base_name}_converted_pdf.pdf"
            output_path = os.path.join(output_dir, output_name)

            self.ppt_to_pdf(ppt_file, output_path)
            created_pdfs.append(output_path)

        if progress_callback:
            progress_callback("All files converted!", 100)

        return created_pdfs

    def cleanup_temp_files(self) -> None:
        for pdf in self.temp_pdfs:
            try:
                os.remove(pdf)
            except OSError:
                pass

        if self.temp_dir:
            try:
                os.rmdir(self.temp_dir)
            except OSError:
                pass

        self.temp_pdfs = []
        self.temp_dir = None
