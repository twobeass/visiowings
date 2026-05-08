"""Lightweight logging setup used by the visiowings CLI.

We keep the dependency footprint minimal by sticking to the stdlib
``logging`` module. Output is colorized only when the stream is a TTY and
``NO_COLOR`` is not set (https://no-color.org/).
"""

from __future__ import annotations

import logging
import os
import sys
from typing import TextIO

_RESET = "\x1b[0m"
_LEVEL_STYLES: dict[int, str] = {
    logging.DEBUG: "\x1b[2m",      # dim
    logging.INFO: "",
    logging.WARNING: "\x1b[33m",   # yellow
    logging.ERROR: "\x1b[31m",     # red
    logging.CRITICAL: "\x1b[1;31m",  # bold red
}


def _supports_color(stream: TextIO) -> bool:
    if os.environ.get("NO_COLOR"):
        return False
    if os.environ.get("FORCE_COLOR"):
        return True
    return getattr(stream, "isatty", lambda: False)()


class ColoringFormatter(logging.Formatter):
    """Formatter that colors the level name when the destination is a TTY."""

    def __init__(self, *args: object, color: bool = True, **kwargs: object) -> None:
        super().__init__(*args, **kwargs)  # type: ignore[arg-type]
        self.color = color

    def format(self, record: logging.LogRecord) -> str:
        rendered = super().format(record)
        if not self.color:
            return rendered
        style = _LEVEL_STYLES.get(record.levelno, "")
        if not style:
            return rendered
        return f"{style}{rendered}{_RESET}"


def setup_logging(
    *,
    debug: bool = False,
    stream: TextIO | None = None,
) -> None:
    """Configure the root logger for CLI output.

    Idempotent: calling this multiple times re-uses (and reconfigures) the
    single handler we install on the root logger.
    """

    stream = stream or sys.stderr
    level = logging.DEBUG if debug else logging.INFO

    fmt = "%(message)s" if not debug else "[%(levelname)s] %(name)s: %(message)s"
    formatter = ColoringFormatter(fmt=fmt, color=_supports_color(stream))

    root = logging.getLogger()
    # Reset any prior handlers so calling setup_logging twice does not
    # accumulate them (e.g. when --debug toggles between invocations in tests).
    for handler in list(root.handlers):
        root.removeHandler(handler)

    handler = logging.StreamHandler(stream)
    handler.setFormatter(formatter)
    handler.setLevel(level)
    root.addHandler(handler)
    root.setLevel(level)


__all__ = ["ColoringFormatter", "setup_logging"]
