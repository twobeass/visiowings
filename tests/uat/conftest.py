"""pytest configuration for the in-tree UAT suite.

The UAT suite exercises a real `visiowings` CLI against a live Visio
instance — distinct from the unit tests in `tests/test_*.py`, which run
against COM mocks and don't need Visio.

Responsibilities:
    - Register UAT markers (delegates to ``tests.uat.markers``).
    - Provide session/function fixtures for the repo root, fixtures dir, COM apps.
    - Auto-skip based on environment (no Office app, no docker, no network, ...).
    - Session-finalize: kill any Office zombies left behind.

Importantly: this conftest is scoped under ``tests/uat/`` and does NOT
affect the unit tests in the parent ``tests/`` directory.
"""

from __future__ import annotations

import os
import shutil
import sys
from pathlib import Path

import pytest

# Repository root: <repo>/tests/uat/conftest.py → parents[2]
REPO_ROOT = Path(__file__).resolve().parents[2]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

# The UAT test modules import ``winreg`` / ``pywin32`` at top level. On
# non-Windows runners (Linux/macOS CI) even *collecting* them raises
# ``ModuleNotFoundError: winreg``, so we short-circuit collection at the
# conftest level instead of relying on per-test skips.
collect_ignore_glob = ["test_*.py"] if sys.platform != "win32" else []

from tests.uat import markers as _markers  # noqa: E402
from tests.uat.com_helpers.process import kill_zombies  # noqa: E402

_PYWIN32_MODULE_NAMES = ("win32com.client", "win32com", "pythoncom", "pywintypes")


@pytest.fixture(autouse=True)
def _restore_real_pywin32():
    """Drop the parent conftest's MagicMock stubs for the duration of each UAT test.

    The repo-wide ``tests/conftest.py`` installs ``MagicMock`` stand-ins for
    ``win32com`` / ``pythoncom`` / ``pywintypes`` so unit tests can run on
    Linux. The UAT suite needs the real COM bindings; we pop the stubs
    around each test and reinstate them afterwards so subsequent unit
    tests still see the mocks they expect.
    """
    if sys.platform != "win32":
        yield
        return
    saved: dict[str, object] = {}
    for name in _PYWIN32_MODULE_NAMES:
        if name in sys.modules:
            saved[name] = sys.modules.pop(name)
    try:
        yield
    finally:
        # Drop any real pywin32 module that got imported during this test,
        # then put the stubs back. Order matters: pop real modules first so
        # ``sys.modules[name] = stub`` is a clean replacement.
        for name in _PYWIN32_MODULE_NAMES:
            if name not in saved:
                sys.modules.pop(name, None)
        for name, stub in saved.items():
            sys.modules[name] = stub


# ---------------------------------------------------------------------------
# pytest hooks
# ---------------------------------------------------------------------------


def pytest_configure(config):
    _markers.register(config)


def pytest_collection_modifyitems(config, items):
    """Apply environment-based skips so missing apps don't fail the run."""
    office_state = _detect_office_cached()
    docker_available = shutil.which("docker") is not None
    network_available = _probe_network()

    for item in items:
        # @requires_office("app")
        for mark in item.iter_markers(name="requires_office"):
            app = (mark.args[0] if mark.args else "").lower()
            if app and not office_state.get(app, False):
                item.add_marker(pytest.mark.skip(reason=f"Office app not installed: {app}"))

        # @requires_docker
        if list(item.iter_markers(name="requires_docker")) and not docker_available:
            item.add_marker(pytest.mark.skip(reason="docker not on PATH"))

        # @requires_network
        if list(item.iter_markers(name="requires_network")) and not network_available:
            item.add_marker(pytest.mark.skip(reason="no network connectivity"))

        # @windows_only
        if list(item.iter_markers(name="windows_only")) and sys.platform != "win32":
            item.add_marker(pytest.mark.skip(reason="windows only"))

        # @not_yet_implemented
        for mark in item.iter_markers(name="not_yet_implemented"):
            iter_label = mark.args[0] if mark.args else "later"
            item.add_marker(pytest.mark.skip(reason=f"NOT YET IMPLEMENTED ({iter_label})"))


def pytest_sessionfinish(session, exitstatus):
    """Best-effort cleanup of any Office zombies left by THIS test session.

    Excludes ``visio.exe`` deliberately: tests that need an open document
    (§C2/§D1/§D2/§E1/§E4) attach to a user-opened Visio via
    ``GetActiveObject`` rather than spawning their own instance. Killing
    every Visio at session end would tear down the user's session and
    force them to re-open the fixture for the next run.
    """
    try:
        kill_zombies(
            names=("excel.exe", "winword.exe", "msaccess.exe", "powerpnt.exe", "outlook.exe")
        )
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Cached environment probes
# ---------------------------------------------------------------------------

_office_cache: dict[str, bool] | None = None


def _detect_office_cached() -> dict[str, bool]:
    global _office_cache
    if _office_cache is not None:
        return _office_cache
    try:
        from tests.uat.setup.office_detect import detect

        _office_cache = detect()
    except Exception:
        _office_cache = {}
    return _office_cache


def _probe_network() -> bool:
    if os.environ.get("UAT_OFFLINE") == "1":
        return False
    import socket

    try:
        socket.create_connection(("github.com", 443), timeout=3).close()
        return True
    except OSError:
        return False


# ---------------------------------------------------------------------------
# Path fixtures
# ---------------------------------------------------------------------------


@pytest.fixture(scope="session")
def workspace_root() -> Path:
    """Repository root (same as visiowings_repo in this layout)."""
    return REPO_ROOT


@pytest.fixture(scope="session")
def visiowings_repo() -> Path:
    """Path to the visiowings repo root.

    In the in-tree layout this is the repository root itself (the
    ``visiowings/`` package is a subdirectory of it). Kept as a named
    fixture so migrated test bodies remain unchanged.
    """
    return REPO_ROOT


@pytest.fixture(scope="session")
def fixtures_dir(workspace_root) -> Path:
    """Path to the fixtures directory.

    Fixture generation is a separate step (``python -m
    tests.uat.setup.fixture_factory``) — launching Visio and saving a
    .vsdm easily exceeds the 120s pytest-timeout, so we don't fold it
    into a fixture. If the manifest is missing, dependent tests skip
    with a clean reason.
    """
    d = workspace_root / "fixtures"
    if not (d / "manifest.json").exists():
        pytest.skip(
            "fixtures not generated — run `python -m tests.uat.setup.fixture_factory` to build them"
        )
    return d


# ---------------------------------------------------------------------------
# COM app fixtures (function-scoped: each test gets a fresh instance)
# ---------------------------------------------------------------------------


@pytest.fixture
def visio_app():
    from tests.uat.com_helpers.visio import VisioContext

    with VisioContext() as app:
        yield app
