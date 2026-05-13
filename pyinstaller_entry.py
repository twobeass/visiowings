"""PyInstaller entry point for the standalone ``visiowings.exe``.

PyInstaller invokes its analysis script as a top-level module — it has no
parent package — so the relative imports inside ``visiowings/cli.py``
(``from . import __version__`` etc.) explode at import time with
``ImportError: attempted relative import with no known parent package``.

This wrapper sidesteps the issue by importing ``visiowings.cli`` through
the regular package path, then handing control off to its ``main()``.
``visiowings.spec`` references this file as the analysis entry.
"""

from __future__ import annotations

import sys

from visiowings.cli import main

if __name__ == "__main__":
    sys.exit(main() or 0)
