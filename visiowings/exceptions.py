"""Typed exception hierarchy for visiowings.

The CLI catches :class:`VisiowingsError` at the top level and reports the
``message`` attribute to the user without a traceback (unless ``--debug``
is set). Internal code should raise the most specific subclass available.
"""

from __future__ import annotations


class VisiowingsError(Exception):
    """Base class for all expected, user-facing errors."""

    #: A short, user-friendly message. When ``str(exc)`` would print the
    #: full repr, override this in subclasses to keep error output clean.
    message: str = ""

    def __init__(self, message: str = "", *args: object) -> None:
        super().__init__(message, *args)
        self.message = message or self.__class__.__doc__ or ""


class VisioNotRunningError(VisiowingsError):
    """Visio is not running or the COM object cannot be reached.

    Hint: Open Visio first, then re-run the command.
    """


class DocumentNotFoundError(VisiowingsError):
    """The requested ``.vsdm``/``.vsdx`` is not currently open in Visio."""

    def __init__(self, requested: str, available: list[str] | None = None) -> None:
        self.requested = requested
        self.available = list(available or [])
        msg = f"Document not found in Visio: {requested!r}"
        if self.available:
            msg += "\nOpen documents:\n  - " + "\n  - ".join(self.available)
        else:
            msg += "\nNo documents are currently open in Visio."
        super().__init__(msg)


class UnsupportedEncodingError(VisiowingsError):
    """The user requested an encoding that Python's codecs registry does not know."""

    def __init__(self, name: str) -> None:
        self.name = name
        super().__init__(
            f"Unsupported encoding: {name!r}. Try a Windows codepage like "
            f"'cp1252', 'cp1251', 'cp932', 'cp936', or 'cp949'."
        )


class EncodingIncompatibilityError(VisiowingsError):
    """A VBA file contains characters the target codepage cannot represent.

    Raised during import when a `.bas` / `.cls` / `.frm` body contains
    code points outside the document's resolved codepage (e.g. emoji or
    CJK text in a cp1252 document). Re-encoding with ``errors="replace"``
    would silently corrupt the source, and Visio's ``VBComponents.Import``
    behaves erratically on partially-encodable temp files — so we refuse
    the import upfront and surface the offending characters to the user.
    """

    def __init__(self, file: str, codepage: str, sample_chars: list[str]) -> None:
        self.file = file
        self.codepage = codepage
        self.sample_chars = list(sample_chars)
        sample = ", ".join(f"{c!r} (U+{ord(c):04X})" for c in self.sample_chars[:5])
        super().__init__(
            f"{file}: contains characters not representable in {codepage}: {sample}. "
            f"Pass `--codepage cp65001` (UTF-8) or remove the offending characters; "
            f"the original module in Visio was NOT modified."
        )


class COMConnectionError(VisiowingsError):
    """The COM connection to Visio dropped and could not be re-established."""

    def __init__(self, attempts: int, last_error: BaseException | None = None) -> None:
        self.attempts = attempts
        self.last_error = last_error
        suffix = f" Last error: {last_error!r}" if last_error else ""
        super().__init__(
            f"Could not re-establish COM connection to Visio after {attempts} attempt(s).{suffix}"
        )


class VBAImportError(VisiowingsError):
    """Importing a VBA module file into Visio failed."""

    def __init__(self, file: str, reason: str) -> None:
        self.file = file
        self.reason = reason
        super().__init__(f"Could not import {file}: {reason}")


class InvalidVisioFileError(VisiowingsError):
    """The path supplied to ``visiowings`` does not look like a Visio file."""

    SUPPORTED_SUFFIXES: tuple[str, ...] = (
        ".vsd",
        ".vsdx",
        ".vsdm",
        ".vstx",
        ".vstm",
        ".vssm",
        ".vssx",
    )

    def __init__(self, path: str) -> None:
        self.path = path
        super().__init__(
            f"Not a Visio file: {path}. Supported extensions: {', '.join(self.SUPPORTED_SUFFIXES)}"
        )


__all__ = [
    "COMConnectionError",
    "DocumentNotFoundError",
    "EncodingIncompatibilityError",
    "InvalidVisioFileError",
    "UnsupportedEncodingError",
    "VBAImportError",
    "VisioNotRunningError",
    "VisiowingsError",
]
