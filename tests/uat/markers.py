"""pytest marker registration and shared constants.

The UAT suite uses strict markers (configured in pyproject.toml) — every
marker must be declared here, otherwise pytest fails the collection.
"""

from __future__ import annotations

from collections.abc import Iterable

MARKERS = [
    ("section(id, ref='')", "UAT section identifier and optional doc anchor"),
    ("requires_office(app)", "Skip if the named Office app is not installed"),
    ("requires_python_version(spec)", "Skip if interpreter doesn't match PEP 440 spec"),
    ("requires_docker", "Skip if docker is not on PATH"),
    ("requires_network", "Skip if no network access (mainly for API-backed steps)"),
    ("manual_signoff(reason)", "Set up the precondition, then skip and surface in report"),
    ("windows_only", "Skip when not running on Windows"),
    ("not_yet_implemented(iter)", "Marks a test as a placeholder for a later iteration"),
]


def register(config) -> None:
    for spec, desc in MARKERS:
        config.addinivalue_line("markers", f"{spec}: {desc}")


def iter_marker_names() -> Iterable[str]:
    for spec, _ in MARKERS:
        yield spec.split("(", 1)[0].strip()
