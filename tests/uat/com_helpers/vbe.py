"""Shared VBA Project / VBE helpers.

The VBComponent types we care about (from ``vbext_ComponentType``):

    1 = vbext_ct_StdModule
    2 = vbext_ct_ClassModule
    3 = vbext_ct_MSForm
   11 = vbext_ct_ActiveXDesigner
  100 = vbext_ct_Document  (ThisDocument, ThisWorkbook, ...)

If Trust Center "Access to VBA project object model" is not granted,
accessing ``doc.VBProject`` raises pywintypes.com_error with HResult
``0x800A03EC``. We translate that to ``VBOMAccessDeniedError`` so callers
(and bootstrap) can react.
"""

from __future__ import annotations

VBOM_DENIED_HRESULT = -2146827284  # 0x800A03EC as signed int32

VBEXT_CT_STDMODULE = 1
VBEXT_CT_CLASSMODULE = 2
VBEXT_CT_MSFORM = 3
VBEXT_CT_DOCUMENT = 100


class VBOMAccessDeniedError(RuntimeError):
    """Raised when AccessVBOM is not effective for the running Office app."""


def _vbproject(doc):
    try:
        return doc.VBProject
    except Exception as exc:
        msg = str(exc)
        if "0x800A03EC" in msg or str(VBOM_DENIED_HRESULT) in msg:
            raise VBOMAccessDeniedError(
                "Trust Center 'AccessVBOM' is not effective. Run "
                "tests/uat/setup/trust_center.py and restart the Office app."
            ) from exc
        raise


def list_modules(doc) -> list[str]:
    proj = _vbproject(doc)
    return [c.Name for c in proj.VBComponents]


def read_module_code(doc, module_name: str) -> str:
    proj = _vbproject(doc)
    comp = proj.VBComponents(module_name)
    code = comp.CodeModule
    count = code.CountOfLines
    if count <= 0:
        return ""
    return code.Lines(1, count)


def _strip_leading_option_explicit(code: str) -> str:
    """Drop a single leading ``Option Explicit`` line (and the blank line
    after it, if any). Used before ``AddFromString`` when the VBE has
    already auto-injected ``Option Explicit`` because „Require Variable
    Declaration" is enabled.
    """
    lines = code.splitlines(keepends=True)
    stripped = False
    out: list[str] = []
    for ln in lines:
        if not stripped and ln.strip().lower() == "option explicit":
            stripped = True
            continue
        if stripped and not out and ln.strip() == "":
            continue
        out.append(ln)
    return "".join(out)


def _add_code_dedup_option_explicit(comp, code: str) -> None:
    """``AddFromString`` while de-duplicating ``Option Explicit`` against
    whatever the VBE already pre-seeded into the freshly created module.
    """
    existing = ""
    if comp.CodeModule.CountOfLines > 0:
        existing = comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)
    if "option explicit" in existing.lower():
        code = _strip_leading_option_explicit(code)
    comp.CodeModule.AddFromString(code)


def add_std_module(doc, name: str, code: str):
    proj = _vbproject(doc)
    comp = proj.VBComponents.Add(VBEXT_CT_STDMODULE)
    comp.Name = name
    if code:
        _add_code_dedup_option_explicit(comp, code)
    return comp


def add_class_module(doc, name: str, code: str):
    proj = _vbproject(doc)
    comp = proj.VBComponents.Add(VBEXT_CT_CLASSMODULE)
    comp.Name = name
    if code:
        _add_code_dedup_option_explicit(comp, code)
    return comp


def add_userform(doc, name: str, code: str | None = None):
    proj = _vbproject(doc)
    comp = proj.VBComponents.Add(VBEXT_CT_MSFORM)
    comp.Name = name
    if code:
        _add_code_dedup_option_explicit(comp, code)
    return comp


def remove_module(doc, module_name: str) -> bool:
    proj = _vbproject(doc)
    try:
        comp = proj.VBComponents(module_name)
    except Exception:
        return False
    proj.VBComponents.Remove(comp)
    return True
