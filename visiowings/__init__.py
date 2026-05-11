"""visiowings - VBA Editor for Microsoft Visio.

Top-level imports of the COM-touching classes are intentionally lazy:
``pywin32`` is only installable on Windows, so an eager
``from .file_watcher import VBAWatcher`` would crash the entry point on
Linux/macOS — including ``visiowings --help`` and
``visiowings init --non-interactive``, which don't need COM at all.

Consumers can still use ``from visiowings import VBAWatcher``; the class
is materialised on first attribute access via :pep:`562`'s
``__getattr__``.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

__version__ = "0.6.1"
__author__ = "twobeass"

__all__ = ["VBAWatcher", "VisioVBAExporter", "VisioVBAImporter"]

_LAZY_EXPORTS = {
    "VBAWatcher": ("file_watcher", "VBAWatcher"),
    "VisioVBAExporter": ("vba_export", "VisioVBAExporter"),
    "VisioVBAImporter": ("vba_import", "VisioVBAImporter"),
}


def __getattr__(name: str) -> Any:
    target = _LAZY_EXPORTS.get(name)
    if target is None:
        raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
    module_name, attr = target
    from importlib import import_module

    module = import_module(f"{__name__}.{module_name}")
    return getattr(module, attr)


if TYPE_CHECKING:  # pragma: no cover
    # `TC004` flags these as "used outside TYPE_CHECKING" because of the
    # `__all__` / `_LAZY_EXPORTS` references — but at runtime those go
    # through `__getattr__` and never need the static import. Keeping the
    # imports in the type-checking block is what lets mypy + IDEs resolve
    # the symbols without pulling pywin32 in at import time.
    from .file_watcher import VBAWatcher
    from .vba_export import VisioVBAExporter
    from .vba_import import VisioVBAImporter
