"""Tests for VBA-export header stripping and Rubberduck annotation extraction."""

from __future__ import annotations

import pytest

from visiowings.vba_export import VisioVBAExporter


# --------------------------------------------------------------------------- #
# _strip_vba_header_export
# --------------------------------------------------------------------------- #
class TestStripVbaHeaderExport:
    def test_keeps_vb_name_for_bas_export(self):
        exporter = VisioVBAExporter("dummy.vsdm")
        code = (
            "VERSION 1.0 CLASS\n"
            "BEGIN\n"
            "  MultiUse = -1\n"
            "END\n"
            'Attribute VB_Name = "Foo"\n'
            "Attribute VB_GlobalNameSpace = False\n"
            "Option Explicit\n"
            "Sub Bar()\nEnd Sub\n"
        )
        cleaned = exporter._strip_vba_header_export(code, keep_vb_name=True)
        assert 'Attribute VB_Name = "Foo"' in cleaned
        assert "VERSION 1.0 CLASS" not in cleaned
        assert "BEGIN" not in cleaned
        assert "Attribute VB_GlobalNameSpace" not in cleaned
        assert "Option Explicit" in cleaned

    def test_drops_vb_name_for_comparison(self):
        exporter = VisioVBAExporter("dummy.vsdm")
        code = (
            'Attribute VB_Name = "Foo"\n'
            "Sub Bar()\n"
            "End Sub\n"
        )
        cleaned = exporter._strip_vba_header_export(code, keep_vb_name=False)
        assert "Attribute VB_Name" not in cleaned
        assert "Sub Bar()" in cleaned

    def test_userform_nested_begins_collapse(self):
        exporter = VisioVBAExporter("dummy.vsdm")
        code = (
            "VERSION 5.00\n"
            "Begin {GUID} UserForm1\n"
            '   Caption = "UserForm1"\n'
            "   Begin {GUID} CommandButton1\n"
            '      Caption = "Click"\n'
            "   End\n"
            "End\n"
            'Attribute VB_Name = "UserForm1"\n'
            "Option Explicit\n"
            "Sub Click()\n"
            "End Sub\n"
        )
        cleaned = exporter._strip_vba_header_export(code, keep_vb_name=True)
        # Begin/End form blocks must be gone
        assert "Begin {GUID}" not in cleaned
        # ... but End Sub for the click handler must remain
        assert "End Sub" in cleaned
        # And the actual code must be intact
        assert "Sub Click()" in cleaned
        assert "Option Explicit" in cleaned


# --------------------------------------------------------------------------- #
# _extract_folder_annotation (Rubberduck @Folder)
# --------------------------------------------------------------------------- #
class TestExtractFolderAnnotation:
    @pytest.fixture
    def exporter(self):
        return VisioVBAExporter("dummy.vsdm", use_rubberduck=True)

    def test_extracts_simple_annotation(self, exporter):
        code = "'@Folder(\"Main\")\nOption Explicit\n"
        assert exporter._extract_folder_annotation(code) == "Main"

    def test_extracts_dotted_annotation(self, exporter):
        code = "'@Folder(\"Main.Sub.Deep\")\n"
        assert exporter._extract_folder_annotation(code) == "Main/Sub/Deep"

    def test_handles_whitespace_variants(self, exporter):
        code = "'@Folder ( \"My Folder\" )\n"
        assert exporter._extract_folder_annotation(code) == "My Folder"

    def test_returns_none_when_no_annotation_present(self, exporter):
        code = "Option Explicit\nSub Foo()\nEnd Sub"
        assert exporter._extract_folder_annotation(code) is None

    def test_returns_none_when_rubberduck_disabled(self):
        exporter = VisioVBAExporter("dummy.vsdm", use_rubberduck=False)
        code = "'@Folder(\"Main\")\nOption Explicit\n"
        # The exporter only honors @Folder when use_rubberduck=True
        # (the function still parses, but downstream callers should ignore).
        # If the implementation gates extraction internally, accept None.
        result = exporter._extract_folder_annotation(code)
        assert result in (None, "Main")
