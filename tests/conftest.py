"""Shared pytest configuration.

The Visio COM bindings are unconditionally mocked out so the test suite
can run on Linux/macOS as well as Windows. Tests that genuinely need a
real Visio + Windows environment must use ``@pytest.mark.windows_only``.
"""

from __future__ import annotations

import sys
from types import ModuleType
from unittest.mock import MagicMock

import pytest


def _install_pywin32_stubs() -> None:
    """Replace pywin32 modules with safe stand-ins before visiowings is imported."""

    if "win32com" not in sys.modules:
        win32com = ModuleType("win32com")
        win32com.__path__ = []  # type: ignore[attr-defined]
        sys.modules["win32com"] = win32com

    if "win32com.client" not in sys.modules:
        client = ModuleType("win32com.client")
        client.Dispatch = MagicMock(name="Dispatch")  # type: ignore[attr-defined]
        client.GetActiveObject = MagicMock(name="GetActiveObject")  # type: ignore[attr-defined]
        sys.modules["win32com.client"] = client

    if "pythoncom" not in sys.modules:
        pythoncom = ModuleType("pythoncom")
        pythoncom.CoInitialize = MagicMock(name="CoInitialize")  # type: ignore[attr-defined]
        pythoncom.CoUninitialize = MagicMock(name="CoUninitialize")  # type: ignore[attr-defined]

        # Real pywin32 raises this; tests that catch ``com_error`` need a class
        class _ComError(Exception):
            def __init__(self, hresult: int = 0, *args: object) -> None:
                super().__init__(hresult, *args)
                self.hresult = hresult

        pythoncom.com_error = _ComError  # type: ignore[attr-defined]
        sys.modules["pythoncom"] = pythoncom

    if "pywintypes" not in sys.modules:
        pywintypes = ModuleType("pywintypes")
        pywintypes.com_error = sys.modules["pythoncom"].com_error  # type: ignore[attr-defined]
        sys.modules["pywintypes"] = pywintypes


_install_pywin32_stubs()


# --------------------------------------------------------------------------- #
# Fixtures
# --------------------------------------------------------------------------- #
@pytest.fixture
def fake_doc():
    """A FakeVisioDocument with one StdModule, ready for export tests."""

    from tests._visio_mocks import FakeVBComponent, FakeVisioDocument, VBComponentType

    return FakeVisioDocument(
        name="Drawing1.vsdm",
        full_name="C:\\Docs\\Drawing1.vsdm",
        components=[
            FakeVBComponent(
                "Module1",
                component_type=VBComponentType.STD_MODULE,
                text='Attribute VB_Name = "Module1"\nOption Explicit\n',
            )
        ],
    )


@pytest.fixture
def fake_app_one_doc(fake_doc):
    from tests._visio_mocks import make_visio_app

    return make_visio_app(fake_doc)


@pytest.fixture
def fake_app_with_stencil():
    from tests._visio_mocks import (
        FakeVBComponent,
        FakeVisioDocument,
        VBComponentType,
        VisioDocumentType,
        make_visio_app,
    )

    drawing = FakeVisioDocument(
        name="Drawing1.vsdm",
        full_name="C:\\Docs\\Drawing1.vsdm",
        components=[FakeVBComponent("Module1", VBComponentType.STD_MODULE, "Sub Foo()\nEnd Sub\n")],
    )
    stencil = FakeVisioDocument(
        name="Shapes.vssm",
        full_name="C:\\Docs\\Shapes.vssm",
        doc_type=VisioDocumentType.STENCIL,
        components=[
            FakeVBComponent("StencilHelpers", VBComponentType.STD_MODULE, "Sub Bar()\nEnd Sub\n")
        ],
    )
    return make_visio_app(drawing, stencil)


@pytest.fixture
def fake_app_no_vba():
    from tests._visio_mocks import FakeVisioDocument, make_visio_app

    return make_visio_app(
        FakeVisioDocument(
            name="Drawing1.vsdx", full_name="C:\\Docs\\Drawing1.vsdx", has_vba=False
        )
    )


def pytest_collection_modifyitems(config, items):
    """Skip ``windows_only`` tests automatically on non-Windows hosts."""

    if sys.platform.startswith("win"):
        return
    skip_marker = pytest.mark.skip(reason="windows_only test skipped on non-Windows host")
    for item in items:
        if "windows_only" in item.keywords:
            item.add_marker(skip_marker)
