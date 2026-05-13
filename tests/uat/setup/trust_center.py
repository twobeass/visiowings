"""Trust Center registry writer — enables ``AccessVBOM`` for every installed
Office application across known Office major versions.

Equivalent to ticking *"Trust access to the VBA project object model"* in
File → Options → Trust Center → Macro Settings, but for every app at once.

Idempotent: writing the same DWORD twice is a no-op.
"""

from __future__ import annotations

import winreg

OFFICE_VERSIONS = ["16.0", "15.0"]  # 365/2019/2021/2024 = 16.0, 2013 = 15.0
APPS = ["Visio", "Excel", "Word", "Access", "PowerPoint", "Outlook"]


def enable_vbom_access() -> list[tuple[str, str]]:
    """Set AccessVBOM=1 and VBAWarnings=1 for every (app, version) pair.

    Returns the list of (app, version) tuples that were successfully touched.
    Silently skips apps not installed in a given Office version.
    """
    touched: list[tuple[str, str]] = []
    for ver in OFFICE_VERSIONS:
        for app in APPS:
            key_path = rf"Software\Microsoft\Office\{ver}\{app}\Security"
            try:
                key = winreg.CreateKeyEx(
                    winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE
                )
                try:
                    winreg.SetValueEx(key, "AccessVBOM", 0, winreg.REG_DWORD, 1)
                    winreg.SetValueEx(key, "VBAWarnings", 0, winreg.REG_DWORD, 1)
                    touched.append((app, ver))
                finally:
                    winreg.CloseKey(key)
            except OSError:
                continue
    return touched


def read_vbom_state() -> dict[tuple[str, str], dict[str, int]]:
    """Diagnostic read-back. Returns {(app, ver): {AccessVBOM: x, VBAWarnings: y}}."""
    state: dict[tuple[str, str], dict[str, int]] = {}
    for ver in OFFICE_VERSIONS:
        for app in APPS:
            key_path = rf"Software\Microsoft\Office\{ver}\{app}\Security"
            try:
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as key:
                    values: dict[str, int] = {}
                    for name in ("AccessVBOM", "VBAWarnings"):
                        try:
                            v, _ = winreg.QueryValueEx(key, name)
                            values[name] = int(v)
                        except OSError:
                            pass
                    if values:
                        state[(app, ver)] = values
            except OSError:
                continue
    return state


if __name__ == "__main__":
    touched = enable_vbom_access()
    print(f"AccessVBOM set on {len(touched)} app/version pair(s):")
    for app, ver in touched:
        print(f"  - {app} {ver}")
