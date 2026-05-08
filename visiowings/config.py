"""Per-project configuration handling.

A ``.visiowings.toml`` file in the working directory lets users skip the
explicit ``--file`` / ``--output`` / ``--codepage`` / ``--bidirectional``
flags every time. The schema is intentionally tiny:

.. code-block:: toml

    file = "drawings/main.vsdm"
    output = "vba"
    codepage = "cp1252"
    bidirectional = true
    rubberduck = false

The file is created by ``visiowings init`` and consumed by every
subcommand that supports the relevant flags.
"""

from __future__ import annotations

import sys
from dataclasses import dataclass, field
from pathlib import Path

if sys.version_info >= (3, 11):
    import tomllib
else:  # pragma: no cover - we declare requires-python = '>=3.10'
    import tomli as tomllib  # type: ignore[import-not-found]


CONFIG_FILENAME = ".visiowings.toml"


@dataclass
class VisiowingsConfig:
    """In-memory representation of ``.visiowings.toml``."""

    file: str | None = None
    output: str | None = None
    input: str | None = None
    codepage: str | None = None
    bidirectional: bool = False
    rubberduck: bool = False
    sync_delete_modules: bool = False
    force: bool = False
    extras: dict[str, object] = field(default_factory=dict)

    def to_toml(self) -> str:
        """Serialise the config to a TOML string suitable for writing to disk."""

        lines: list[str] = [
            "# visiowings project config — created by `visiowings init`",
            "# Edit this file (or delete it) to change defaults.",
            "",
        ]
        if self.file is not None:
            lines.append(f'file = "{self.file}"')
        if self.output is not None:
            lines.append(f'output = "{self.output}"')
        if self.input is not None:
            lines.append(f'input = "{self.input}"')
        if self.codepage is not None:
            lines.append(f'codepage = "{self.codepage}"')
        lines.append(f"bidirectional = {str(self.bidirectional).lower()}")
        lines.append(f"rubberduck = {str(self.rubberduck).lower()}")
        lines.append(f"sync_delete_modules = {str(self.sync_delete_modules).lower()}")
        lines.append(f"force = {str(self.force).lower()}")
        lines.append("")
        return "\n".join(lines)


def find_config(start: Path | None = None) -> Path | None:
    """Walk up from ``start`` looking for a ``.visiowings.toml`` file."""

    if start is None:
        start = Path.cwd()
    for parent in [start.resolve(), *start.resolve().parents]:
        candidate = parent / CONFIG_FILENAME
        if candidate.is_file():
            return candidate
    return None


def load_config(path: Path | None = None) -> VisiowingsConfig:
    """Load a ``.visiowings.toml`` from ``path`` (or auto-discover)."""

    if path is None:
        path = find_config()
    if path is None:
        return VisiowingsConfig()

    with path.open("rb") as fh:
        data = tomllib.load(fh)

    known = {
        "file",
        "output",
        "input",
        "codepage",
        "bidirectional",
        "rubberduck",
        "sync_delete_modules",
        "force",
    }
    cfg = VisiowingsConfig()
    for key, value in data.items():
        if key in known:
            setattr(cfg, key, value)
        else:
            cfg.extras[key] = value
    return cfg


def write_config(cfg: VisiowingsConfig, path: Path | None = None) -> Path:
    """Persist ``cfg`` to ``path`` (defaults to ``./visiowings.toml``)."""

    target = (path or (Path.cwd() / CONFIG_FILENAME)).resolve()
    target.write_text(cfg.to_toml(), encoding="utf-8")
    return target


__all__ = [
    "CONFIG_FILENAME",
    "VisiowingsConfig",
    "find_config",
    "load_config",
    "write_config",
]
