"""Generate the minimal Visio fixtures used by the UAT runner.

Creates ``fixtures/sample.vsdm`` — a macro-enabled drawing with one
standard module, one class module, and one user form. Other fixtures
(``sample_stencil.vssm``, ``sample_template.vstm``, ``sample_german.vsdm``)
are stubbed for later iterations.

The manifest stores **code hashes** per module, not file hashes — Visio
embeds non-deterministic IDs so two valid runs produce different bytes.
"""

from __future__ import annotations

import hashlib
import json
from pathlib import Path

from ..com_helpers.vbe import add_class_module, add_std_module, add_userform
from ..com_helpers.visio import VisioContext

BAS_TEMPLATE = """Option Explicit

' Minimal clean standard module used by D2 export/import roundtrip tests.
Public Sub Hello()
    Debug.Print "hello from BasicLogic"
End Sub

Public Function Add(ByVal a As Long, ByVal b As Long) As Long
    Add = a + b
End Function
"""

CLS_TEMPLATE = """Option Explicit

' Simple class module
Private mName As String

Public Property Get Name() As String
    Name = mName
End Property

Public Property Let Name(ByVal value As String)
    mName = value
End Property
"""

FRM_CODE_TEMPLATE = """Option Explicit

Private Sub UserForm_Initialize()
    Me.Caption = "Demo Form"
End Sub
"""


def _hash(s: str) -> str:
    return hashlib.sha256(s.encode("utf-8")).hexdigest()[:16]


def generate_sample_vsdm(output_path: Path) -> dict[str, str]:
    """Create a macro-enabled Visio file with three VBA components.

    Returns a dict of {module_name: code_hash} for the manifest.
    """
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    hashes: dict[str, str] = {}
    with VisioContext() as app:
        doc = app.Documents.Add("")  # blank drawing
        try:
            add_std_module(doc, "BasicLogic", BAS_TEMPLATE)
            hashes["BasicLogic"] = _hash(BAS_TEMPLATE)

            add_class_module(doc, "Person", CLS_TEMPLATE)
            hashes["Person"] = _hash(CLS_TEMPLATE)

            add_userform(doc, "frmDemo", FRM_CODE_TEMPLATE)
            hashes["frmDemo"] = _hash(FRM_CODE_TEMPLATE)

            doc.SaveAs(str(output_path))
        finally:
            try:
                doc.Close()
            except Exception:
                pass

    return hashes


def write_manifest(fixtures_dir: Path, entries: dict[str, dict[str, str]]) -> Path:
    """``entries`` is {fixture_filename: {module_name: code_hash}}."""
    fixtures_dir = Path(fixtures_dir)
    fixtures_dir.mkdir(parents=True, exist_ok=True)
    manifest = fixtures_dir / "manifest.json"
    manifest.write_text(
        json.dumps({"version": 1, "fixtures": entries}, indent=2),
        encoding="utf-8",
    )
    return manifest


def generate_all(fixtures_dir: Path) -> Path:
    fixtures_dir = Path(fixtures_dir)
    fixtures_dir.mkdir(parents=True, exist_ok=True)
    entries: dict[str, dict[str, str]] = {}
    entries["sample.vsdm"] = generate_sample_vsdm(fixtures_dir / "sample.vsdm")
    entries["sample_german.vsdm"] = {"__status__": "deferred"}
    entries["sample_stencil.vssm"] = {"__status__": "deferred"}
    entries["sample_template.vstm"] = {"__status__": "deferred"}
    return write_manifest(fixtures_dir, entries)


if __name__ == "__main__":
    import sys

    # Default fixtures dir: <repo>/fixtures (parents[3] = repo root from
    # tests/uat/setup/fixture_factory.py)
    repo_root = Path(__file__).resolve().parents[3]
    target = Path(sys.argv[1]) if len(sys.argv) > 1 else repo_root / "fixtures"
    out = generate_all(target)
    print(f"manifest: {out}")
