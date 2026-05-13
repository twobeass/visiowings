"""Detect installed Office applications via the Windows registry.

Detection rule (in priority order):
    1. ``HKCR\\<ProgID>\\CLSID`` — proves the COM class is registered.
    2. ``HKLM\\Software\\Microsoft\\Office\\<ver>\\<App>\\InstallRoot\\Path``
       — proves the app payload is on disk for any Office ≥ 2013.

Either signal alone is enough. We avoid actually instantiating any app.
"""

from __future__ import annotations

import winreg

OFFICE_APPS: dict[str, str] = {
    "visio": "Visio.Application",
    "excel": "Excel.Application",
    "word": "Word.Application",
    "access": "Access.Application",
    "powerpoint": "PowerPoint.Application",
    "outlook": "Outlook.Application",
}

APP_REGISTRY_NAMES: dict[str, str] = {
    "visio": "Visio",
    "excel": "Excel",
    "word": "Word",
    "access": "Access",
    "powerpoint": "PowerPoint",
    "outlook": "Outlook",
}

OFFICE_VERSIONS = ("16.0", "15.0")


def _has_progid(prog_id: str) -> bool:
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, rf"{prog_id}\CLSID"):
            return True
    except OSError:
        return False


def _has_install_root(app_name: str) -> bool:
    for ver in OFFICE_VERSIONS:
        for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            try:
                with winreg.OpenKey(
                    hive,
                    rf"Software\Microsoft\Office\{ver}\{app_name}\InstallRoot",
                ) as key:
                    winreg.QueryValueEx(key, "Path")
                    return True
            except OSError:
                continue
    return False


def is_app_available(name: str) -> bool:
    prog_id = OFFICE_APPS.get(name, "")
    app_name = APP_REGISTRY_NAMES.get(name, "")
    if prog_id and _has_progid(prog_id):
        return True
    if app_name and _has_install_root(app_name):
        return True
    return False


def detect() -> dict[str, bool]:
    return {name: is_app_available(name) for name in OFFICE_APPS}


if __name__ == "__main__":
    for name, present in detect().items():
        flag = "OK " if present else "-- "
        print(f"{flag} {name}")
