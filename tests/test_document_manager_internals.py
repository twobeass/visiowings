"""Internal helpers in visiowings.document_manager.

Covers ``sanitize_document_name`` (pure function) and ``VisioDocumentInfo``
construction against the spec-based fakes from ``tests/_visio_mocks.py``.
"""

from __future__ import annotations

import pytest

from tests._visio_mocks import (
    FakeVBComponent,
    FakeVisioDocument,
    VBComponentType,
)
from visiowings.document_manager import (
    VisioDocumentInfo,
    VisioDocumentType,
    sanitize_document_name,
)


# --------------------------------------------------------------------------- #
# sanitize_document_name
# --------------------------------------------------------------------------- #
@pytest.mark.parametrize(
    "raw,expected",
    [
        ("Drawing1.vsdm", "drawing1"),
        ("My Drawing.vsdx", "my_drawing"),
        ("Test Drawing", "test_drawing"),
        ("UPPER.vsdm", "upper"),
        # Only the last `.<ext>` is stripped; intermediate dots are preserved.
        ("file.with.dots.vsdm", "file.with.dots"),
        ("Multi   Spaces.vsdm", "multi_spaces"),
        ("Bad/Path?.vsdm", "bad_path"),
        ("trim___underscores___.vsdm", "trim_underscores"),
        # Edge cases that fall through to the "document" sentinel
        ("", "document"),
        (".", "document"),  # rsplit eats the last "."
        ("/\\?*<>", "document"),  # all chars become "_" then get stripped
        ("a", "a"),
        ("a.b", "a"),
    ],
)
def test_sanitize_document_name(raw, expected):
    assert sanitize_document_name(raw) == expected


def test_sanitize_collapses_consecutive_underscores():
    assert sanitize_document_name("a___b") == "a_b"


def test_sanitize_strips_leading_and_trailing_underscores():
    assert sanitize_document_name("___abc___.vsdm") == "abc"


def test_sanitize_lowercases_unicode():
    # Visio docs from non-ASCII locales should still produce a usable folder
    # name. The function delegates to ``str.lower``, so the result keeps the
    # underlying characters lowercased.
    assert sanitize_document_name("ÄÖÜ.vsdm") == "äöü"


# --------------------------------------------------------------------------- #
# VisioDocumentInfo
# --------------------------------------------------------------------------- #
class TestVisioDocumentInfo:
    def test_basic_attributes_from_fake_doc(self):
        doc = FakeVisioDocument(
            name="Drawing1.vsdm",
            full_name="C:\\Docs\\Drawing1.vsdm",
            doc_type=VisioDocumentType.DRAWING,
            components=[
                FakeVBComponent("Module1", VBComponentType.STD_MODULE, "Sub Foo()\nEnd Sub")
            ],
        )
        info = VisioDocumentInfo(doc)
        assert info.name == "Drawing1.vsdm"
        assert info.full_name == "C:\\Docs\\Drawing1.vsdm"
        assert info.type == VisioDocumentType.DRAWING
        assert info.has_vba is True
        assert info.folder_name == "drawing1"

    def test_has_vba_false_when_no_components(self):
        doc = FakeVisioDocument(name="Empty.vsdm", full_name="C:\\Empty.vsdm", components=[])
        info = VisioDocumentInfo(doc)
        assert info.has_vba is False

    def test_has_vba_false_when_vbproject_missing(self):
        doc = FakeVisioDocument(name="NoVBA.vsdx", full_name="C:\\NoVBA.vsdx", has_vba=False)
        info = VisioDocumentInfo(doc)
        assert info.has_vba is False

    def test_get_type_name_drawing(self):
        doc = FakeVisioDocument(name="d.vsdm", doc_type=VisioDocumentType.DRAWING)
        info = VisioDocumentInfo(doc)
        assert info.get_type_name() == "Drawing"

    def test_get_type_name_stencil(self):
        doc = FakeVisioDocument(
            name="s.vssm",
            doc_type=VisioDocumentType.STENCIL,
            components=[FakeVBComponent("Helpers", text="Sub H()\nEnd Sub")],
        )
        info = VisioDocumentInfo(doc)
        assert info.get_type_name() == "Stencil"

    def test_get_type_name_unknown(self):
        doc = FakeVisioDocument(name="x.tmp", doc_type=999, components=[FakeVBComponent("M")])
        info = VisioDocumentInfo(doc)
        assert info.get_type_name() == "Unknown"

    def test_repr_includes_key_attributes(self):
        doc = FakeVisioDocument(
            name="Drawing1.vsdm",
            components=[FakeVBComponent("M")],
        )
        info = VisioDocumentInfo(doc)
        rendered = repr(info)
        assert "Drawing1.vsdm" in rendered
        assert "Drawing" in rendered
        assert "True" in rendered
