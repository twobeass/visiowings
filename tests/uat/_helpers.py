"""Common helpers for the UAT suite.

Important: CLI invocations target the **current repo's source tree**,
never any globally installed ``visiowings`` binary. The globally installed
script may point at a different checkout or a stale non-editable install,
which would invalidate UAT findings.
"""

from __future__ import annotations

import os
import signal as _signal
import subprocess
import sys
import threading as _threading
import time as _time
from collections.abc import Sequence
from pathlib import Path


def run(
    cmd: Sequence[str], cwd: Path | None = None, timeout: int = 120, env: dict | None = None
) -> subprocess.CompletedProcess:
    """Run a subprocess capturing output. Never shell=True.

    Forces UTF-8 decoding on stdout/stderr so subprocess output from tools
    that emit utf-8 (visiowings reconfigures stdout to utf-8) is decoded
    correctly on cp1252 host shells.
    """
    return subprocess.run(
        list(cmd),
        cwd=str(cwd) if cwd else None,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=timeout,
        env=env,
        check=False,
    )


def _branch_env(repo: Path) -> dict:
    """Force PYTHONPATH to the repo's tree so its in-tree ``visiowings``
    package wins over any system-wide install."""
    env = os.environ.copy()
    existing = env.get("PYTHONPATH", "")
    env["PYTHONPATH"] = str(repo) + (os.pathsep + existing if existing else "")
    return env


def visiowings_cli(repo: Path | None = None) -> list[str]:
    """Return argv prefix to invoke the in-tree visiowings package.

    The package's entry-point per ``pyproject.toml`` is ``visiowings.cli:main``;
    there is no ``__main__.py`` so ``-m visiowings`` would fail. Call the
    submodule explicitly.
    """
    return [sys.executable, "-m", "visiowings.cli"]


def run_branch(
    cmd_tail: Sequence[str],
    repo: Path,
    timeout: int = 120,
    cwd: Path | None = None,
    stdin_input: str | None = None,
) -> subprocess.CompletedProcess:
    """Run with PYTHONPATH pinned to the repo tree.

    ``cwd`` defaults to ``repo`` for the typical case. Pass ``cwd=`` explicitly
    when the test sets up a workdir (e.g. ``.visiowings.toml`` in tmp dir).
    ``stdin_input`` is piped to the subprocess (useful to dismiss interactive
    prompts visiowings emits when modules conflict).
    """
    if stdin_input is None:
        return run(
            list(cmd_tail),
            cwd=(cwd if cwd is not None else repo),
            timeout=timeout,
            env=_branch_env(repo),
        )
    return subprocess.run(
        list(cmd_tail),
        cwd=str(cwd if cwd is not None else repo),
        env=_branch_env(repo),
        input=stdin_input,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=timeout,
        check=False,
    )


def repo_present(repo: Path) -> bool:
    """Sanity check that the repo tree is on disk.

    In the in-tree UAT layout the check is effectively a no-op (the tests
    only run from inside the repo), but it stays here so the migrated test
    bodies do not need to change.
    """
    return repo.exists() and (repo / "visiowings").exists()


# ---------------------------------------------------------------------------
# Background-process harness for `visiowings edit`.
#
# §D3 / §D4 / §D5 / §G5 / §G6 tests need to start `visiowings edit` as a
# watcher, drive its filesystem/COM polling, and shut it down cleanly.
# Wrapping that in a class keeps each test small (~30 LOC).
# ---------------------------------------------------------------------------


class WatcherHandle:
    """Wraps a running ``visiowings edit`` subprocess with stdout
    streaming and a clean shutdown path."""

    def __init__(self, proc, lines: list[str], lock: _threading.Lock):
        self.proc = proc
        self._lines = lines
        self._lock = lock
        self._stopped = False

    @property
    def lines(self) -> list[str]:
        with self._lock:
            return list(self._lines)

    def expect(self, needle: str, timeout: float = 5.0) -> str:
        """Block until ``needle`` appears in any captured stdout line.
        Returns the matching line. Raises ``TimeoutError`` otherwise.
        """
        deadline = _time.monotonic() + timeout
        while _time.monotonic() < deadline:
            for ln in self.lines:
                if needle in ln:
                    return ln
            if self.proc.poll() is not None:
                raise RuntimeError(
                    f"watcher exited (code {self.proc.returncode}) before "
                    f"emitting {needle!r}. Last lines:\n" + "\n".join(self.lines[-15:])
                )
            _time.sleep(0.1)
        raise TimeoutError(
            f"watcher never produced {needle!r} within {timeout}s. "
            f"Last lines:\n" + "\n".join(self.lines[-15:])
        )

    def stop(self, timeout: float = 5.0) -> int:
        """Stop the watcher gracefully. On Windows we send CTRL_BREAK_EVENT
        to the process group (subprocess was created with
        ``CREATE_NEW_PROCESS_GROUP``); on POSIX we send SIGINT.
        Falls back to ``terminate()`` if the process doesn't exit in time.
        Returns the exit code.
        """
        if self._stopped:
            return self.proc.returncode or 0
        self._stopped = True
        if self.proc.poll() is None:
            try:
                if sys.platform == "win32":
                    self.proc.send_signal(_signal.CTRL_BREAK_EVENT)
                else:
                    self.proc.send_signal(_signal.SIGINT)
            except Exception:
                pass
            try:
                self.proc.wait(timeout=timeout)
            except subprocess.TimeoutExpired:
                self.proc.terminate()
                try:
                    self.proc.wait(timeout=2.0)
                except subprocess.TimeoutExpired:
                    self.proc.kill()
                    self.proc.wait(timeout=2.0)
        return self.proc.returncode or 0


def start_watcher(
    visiowings_repo: Path,
    doc_path: str,
    output_dir: Path,
    *extra_args: str,
    timeout: int = 30,
) -> WatcherHandle:
    """Start ``visiowings edit --file <doc> --output <output_dir> <extra>``
    as a background subprocess. stdout/stderr are pumped into a shared list
    by a daemon thread. Caller MUST call ``handle.stop()``.

    On Windows we use ``CREATE_NEW_PROCESS_GROUP`` so we can send
    ``CTRL_BREAK_EVENT`` for a graceful shutdown without killing pytest.
    """
    creation_flags = 0
    if sys.platform == "win32":
        creation_flags = subprocess.CREATE_NEW_PROCESS_GROUP
    # ``-u`` forces unbuffered stdout on the child Python — without it,
    # ``print()`` from visiowings is block-buffered (non-tty pipe) and
    # banners like "Starting Live Synchronization" only flush on exit.
    env = _branch_env(visiowings_repo)
    env["PYTHONUNBUFFERED"] = "1"
    proc = subprocess.Popen(
        [
            sys.executable,
            "-u",
            "-m",
            "visiowings.cli",
            "edit",
            "--file",
            str(doc_path),
            "--output",
            str(output_dir),
            *extra_args,
        ],
        cwd=str(visiowings_repo),
        env=env,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
        bufsize=1,
        creationflags=creation_flags,
    )
    lines: list[str] = []
    lock = _threading.Lock()

    def _pump():
        try:
            for ln in proc.stdout:
                with lock:
                    lines.append(ln.rstrip("\r\n"))
        except Exception:
            pass

    t = _threading.Thread(target=_pump, daemon=True)
    t.start()
    return WatcherHandle(proc, lines, lock)
