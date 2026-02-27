from enum import Enum


class ConversionBackend(Enum):
    POWERPOINT = "powerpoint"
    WPS = "wps"
    LIBREOFFICE = "libreoffice"
    ONLYOFFICE = "onlyoffice"
    KEYNOTE = "keynote"


class ConversionCancelledError(Exception):
    pass


def backend_display_name(backend: ConversionBackend) -> str:
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
