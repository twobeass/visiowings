"""Optional, opt-out PyPI update check.

Runs in a daemon thread on CLI start and prints a single hint to stderr
if a newer release exists. Cached for 24 h. Disabled completely if any
of these is true:

- ``VISIOWINGS_NO_UPDATE_CHECK`` env var is set.
- ``CI`` env var is set (avoid log spam in pipelines).
- ``--no-update-check`` was passed.
- Network resolution fails / times out within 2 seconds.
"""

from __future__ import annotations

import json
import logging
import os
import sys
import threading
import time
from pathlib import Path
from urllib import request as urllib_request
from urllib.error import URLError

from . import __version__

logger = logging.getLogger(__name__)

_PYPI_JSON_URL = "https://pypi.org/pypi/visiowings/json"
_CACHE_DIR = Path.home() / ".cache" / "visiowings"
_CACHE_FILE = _CACHE_DIR / "update_check.json"
_CACHE_TTL_SECONDS = 24 * 60 * 60
_HTTP_TIMEOUT_SECONDS = 2.0


def _is_disabled() -> bool:
    if os.environ.get("VISIOWINGS_NO_UPDATE_CHECK"):
        return True
    if os.environ.get("CI"):
        return True
    return False


def _load_cache() -> dict | None:
    try:
        if not _CACHE_FILE.is_file():
            return None
        if (time.time() - _CACHE_FILE.stat().st_mtime) > _CACHE_TTL_SECONDS:
            return None
        loaded: dict = json.loads(_CACHE_FILE.read_text(encoding="utf-8"))
        return loaded
    except (OSError, json.JSONDecodeError):
        return None


def _save_cache(payload: dict) -> None:
    try:
        _CACHE_DIR.mkdir(parents=True, exist_ok=True)
        _CACHE_FILE.write_text(json.dumps(payload), encoding="utf-8")
    except OSError as e:
        logger.debug("Update-check cache write failed: %s", e)


def _fetch_latest_version() -> str | None:
    try:
        # _PYPI_JSON_URL is a hardcoded https:// constant, not user-controlled.
        with urllib_request.urlopen(  # nosec B310
            _PYPI_JSON_URL, timeout=_HTTP_TIMEOUT_SECONDS
        ) as resp:
            data = json.load(resp)
    except (URLError, TimeoutError, json.JSONDecodeError, OSError) as e:
        logger.debug("PyPI update check failed: %s", e)
        return None
    version = data.get("info", {}).get("version")
    return version if isinstance(version, str) else None


def _parse_version(v: str) -> tuple[int, ...]:
    """Cheap dotted-int parse; handles 0.6.1 / 1.0.0 fine, ignores pre-release tags.

    Stops at the first segment that is not purely numeric, so e.g.
    "0.7.0rc1" parses as (0, 7) and is therefore considered older than
    the released "0.7.0" (which parses as (0, 7, 0)). This matches PEP 440
    ordering closely enough for an update-check hint.
    """

    parts: list[int] = []
    for raw in v.replace("-", ".").replace("_", ".").split("."):
        if not raw.isdigit():
            break
        parts.append(int(raw))
    return tuple(parts)


def _is_outdated(current: str, latest: str) -> bool:
    return _parse_version(latest) > _parse_version(current)


def _check_and_notify() -> None:
    cached = _load_cache()
    latest: str | None
    if cached and "latest" in cached:
        latest = cached["latest"]
    else:
        latest = _fetch_latest_version()
        if latest is not None:
            _save_cache({"latest": latest, "checked_at": time.time()})

    if latest and _is_outdated(__version__, latest):
        sys.stderr.write(
            f"\n💡 visiowings {latest} is available (you have {__version__}).\n"
            f"   Run `pipx upgrade visiowings` (or `pip install -U visiowings`) to update.\n\n"
        )


def schedule_async() -> None:
    """Fire the update check in a daemon thread, then return."""

    if _is_disabled():
        return

    thread = threading.Thread(target=_check_and_notify, name="visiowings-update-check", daemon=True)
    thread.start()


__all__ = ["schedule_async"]
