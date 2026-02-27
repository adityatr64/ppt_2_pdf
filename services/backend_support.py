import os
import shutil
import sys
from typing import List, Optional

from .conversion_types import ConversionBackend, backend_display_name


class BackendSupport:
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

    def find_libreoffice(self) -> Optional[str]:
        if self._libreoffice_path:
            return self._libreoffice_path

        platform_paths = self.LIBREOFFICE_PATHS.get(self._platform_key(), [])
        self._libreoffice_path = self._find_executable(["soffice", "libreoffice"], platform_paths)
        return self._libreoffice_path

    def find_onlyoffice(self) -> Optional[str]:
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
            elif backend == ConversionBackend.LIBREOFFICE and self.find_libreoffice():
                available.append(backend)
            elif backend == ConversionBackend.ONLYOFFICE and self.find_onlyoffice():
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

    def get_active_backend(self) -> ConversionBackend:
        if self._detected_backend:
            return self._detected_backend

        available = self.get_available_backends()
        if not available:
            raise RuntimeError(self.get_install_message())

        self._detected_backend = available[0]
        return self._detected_backend

    def get_active_backend_name(self) -> str:
        return backend_display_name(self.get_active_backend())
