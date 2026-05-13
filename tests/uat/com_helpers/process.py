"""Process utilities — zombie cleanup for Office EXEs."""

from __future__ import annotations

from collections.abc import Iterable

OFFICE_EXES = (
    "visio.exe",
    "excel.exe",
    "winword.exe",
    "msaccess.exe",
    "powerpnt.exe",
    "outlook.exe",
)


def kill_zombies(names: Iterable[str] = OFFICE_EXES, timeout: float = 3.0) -> int:
    """Terminate any matching processes. Returns the count of processes killed.

    Used in COM-helper ``__exit__`` paths and as a session finalizer in
    conftest. Safe to call when no matches exist.
    """
    try:
        import psutil
    except ImportError:
        return 0
    target = {n.lower() for n in names}
    killed = 0
    for proc in psutil.process_iter(["name"]):
        pname = (proc.info.get("name") or "").lower()
        if pname not in target:
            continue
        try:
            proc.terminate()
            try:
                proc.wait(timeout=timeout)
            except psutil.TimeoutExpired:
                proc.kill()
            killed += 1
        except psutil.NoSuchProcess:
            continue
        except Exception:
            continue
    return killed
