"""Spec-based fakes for the Visio COM API surface used by visiowings.

The previous test setup replaced ``win32com``/``pythoncom`` with bare
``MagicMock`` instances. That worked, but a typo in a chained attribute
access (e.g. ``doc.VBProject.VBComponentsX.Add(...)``) would silently
return another MagicMock instead of failing. The fakes here mimic just
enough of Visio's object model to catch those kinds of bugs while staying
easy to instantiate from tests.

Public helpers
--------------
- ``FakeVisioApplication``      — top-level Visio.Application analogue.
- ``FakeVisioDocument``         — Document with VBProject + Type/Language.
- ``FakeVBProject``             — VBProject + VBComponents collection.
- ``FakeVBComponentsCollection``— iteration, ``.Count``, ``.Add``,
                                  ``.Remove``, ``.Import``, item-access.
- ``FakeVBComponent``           — Name/Type + CodeModule.
- ``FakeCodeModule``            — Lines/CountOfLines/AddFromString helpers.
- ``make_visio_app(...)``       — convenience builder for tests.

Codes follow https://learn.microsoft.com/en-us/office/vba/api/vba.vbcomponents
"""

from __future__ import annotations

from typing import Any


class VBComponentType:
    """Mirrors VBA's ``vbext_ComponentType`` enum."""

    STD_MODULE = 1     # vbext_ct_StdModule
    CLASS_MODULE = 2   # vbext_ct_ClassModule
    MS_FORM = 3        # vbext_ct_MSForm
    DOCUMENT = 100     # vbext_ct_Document  (ThisDocument, sheets, ...)


class VisioDocumentType:
    """Mirrors visiowings.document_manager.VisioDocumentType."""

    DRAWING = 1
    STENCIL = 2
    TEMPLATE = 3


# --------------------------------------------------------------------------- #
# CodeModule
# --------------------------------------------------------------------------- #
class FakeCodeModule:
    """Subset of VBA's CodeModule object."""

    __slots__ = ("_text",)

    def __init__(self, text: str = "") -> None:
        self._text = text

    @property
    def CountOfLines(self) -> int:
        if not self._text:
            return 0
        # VBA CountOfLines is the number of lines, not separators.
        return self._text.count("\n") + (0 if self._text.endswith("\n") else 1)

    def Lines(self, start: int = 1, count: int | None = None) -> str:
        lines = self._text.splitlines()
        if count is None:
            count = max(0, len(lines) - start + 1)
        # VBA uses 1-based indexing.
        return "\n".join(lines[start - 1 : start - 1 + count])

    def AddFromString(self, text: str) -> None:
        if self._text and not self._text.endswith("\n"):
            self._text += "\n"
        self._text += text

    def DeleteLines(self, start: int = 1, count: int = 1) -> None:
        lines = self._text.splitlines()
        del lines[start - 1 : start - 1 + count]
        self._text = "\n".join(lines)


# --------------------------------------------------------------------------- #
# VBComponent
# --------------------------------------------------------------------------- #
class FakeVBComponent:
    """A single VBA module/class/form/document module."""

    def __init__(
        self,
        name: str,
        component_type: int = VBComponentType.STD_MODULE,
        text: str = "",
    ) -> None:
        self.Name = name
        self.Type = component_type
        self.CodeModule = FakeCodeModule(text)

    def Export(self, path: str) -> None:  # noqa: N802 - mirrors COM method
        from pathlib import Path

        Path(path).write_text(self.CodeModule.Lines(), encoding="cp1252")


# --------------------------------------------------------------------------- #
# VBComponents collection
# --------------------------------------------------------------------------- #
class FakeVBComponentsCollection:
    """Collection of VBComponents with COM-like semantics."""

    def __init__(self, components: list[FakeVBComponent] | None = None) -> None:
        self._components: list[FakeVBComponent] = list(components or [])

    @property
    def Count(self) -> int:
        return len(self._components)

    def __iter__(self):
        return iter(self._components)

    def __len__(self) -> int:
        return len(self._components)

    def Item(self, index: int | str) -> FakeVBComponent:
        if isinstance(index, int):
            return self._components[index - 1]
        for comp in self._components:
            if comp.Name == index:
                return comp
        raise KeyError(index)

    def Add(self, component_type: int) -> FakeVBComponent:
        name = f"Module{len(self._components) + 1}"
        comp = FakeVBComponent(name, component_type=component_type)
        self._components.append(comp)
        return comp

    def Remove(self, component: FakeVBComponent | str) -> None:
        if isinstance(component, str):
            component = self.Item(component)
        self._components.remove(component)

    def Import(self, file_path: str) -> FakeVBComponent:
        from pathlib import Path

        path = Path(file_path)
        text = path.read_text(encoding="cp1252", errors="replace")

        component_type = VBComponentType.STD_MODULE
        if path.suffix.lower() == ".cls":
            component_type = VBComponentType.CLASS_MODULE
        elif path.suffix.lower() == ".frm":
            component_type = VBComponentType.MS_FORM

        comp = FakeVBComponent(path.stem, component_type=component_type, text=text)
        self._components.append(comp)
        return comp


# --------------------------------------------------------------------------- #
# VBProject
# --------------------------------------------------------------------------- #
class FakeVBProject:
    def __init__(self, components: list[FakeVBComponent] | None = None) -> None:
        self.VBComponents = FakeVBComponentsCollection(components)
        self.Name = "VBAProject"
        self.Protection = 0  # vbext_pp_none


# --------------------------------------------------------------------------- #
# Document
# --------------------------------------------------------------------------- #
class FakeVisioDocument:
    def __init__(
        self,
        name: str,
        full_name: str | None = None,
        doc_type: int = VisioDocumentType.DRAWING,
        language: int = 1033,  # en-US
        components: list[FakeVBComponent] | None = None,
        has_vba: bool = True,
    ) -> None:
        self.Name = name
        self.FullName = full_name or f"C:\\Docs\\{name}"
        self.Type = doc_type
        self.Language = language
        self._has_vba = has_vba
        self.VBProject = FakeVBProject(components) if has_vba else None  # type: ignore[assignment]


class _DocumentsCollection:
    def __init__(self, docs: list[FakeVisioDocument]) -> None:
        self._docs = docs

    def __iter__(self):
        return iter(self._docs)

    def __len__(self) -> int:
        return len(self._docs)

    @property
    def Count(self) -> int:
        return len(self._docs)


class FakeVisioApplication:
    def __init__(self, documents: list[FakeVisioDocument] | None = None) -> None:
        self.Documents = _DocumentsCollection(documents or [])
        self.Name = "Microsoft Visio"
        self.Version = "16.0"


def make_visio_app(*documents: FakeVisioDocument) -> FakeVisioApplication:
    """Convenience builder for tests."""

    return FakeVisioApplication(list(documents))


__all__ = [
    "FakeCodeModule",
    "FakeVBComponent",
    "FakeVBComponentsCollection",
    "FakeVBProject",
    "FakeVisioApplication",
    "FakeVisioDocument",
    "VBComponentType",
    "VisioDocumentType",
    "make_visio_app",
]
