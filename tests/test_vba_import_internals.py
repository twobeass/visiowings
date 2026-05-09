"""Tests for VBA-import header stripping and folder-annotation handling."""

from __future__ import annotations

from pathlib import Path
from unittest.mock import MagicMock

from visiowings.vba_import import VisioVBAImporter


# --------------------------------------------------------------------------- #
# _strip_vba_header
# --------------------------------------------------------------------------- #
class TestStripVbaHeaderImport:
    def test_drops_attributes_when_keep_vb_name_false(self):
        importer = VisioVBAImporter("dummy.vsdm")
        code = (
            'Attribute VB_Name = "Module1"\n'
            "Attribute VB_GlobalNameSpace = False\n"
            "Option Explicit\n"
            "Sub Foo()\n"
            "End Sub\n"
        )
        cleaned = importer._strip_vba_header(code, keep_vb_name=False)
        assert "Attribute VB_Name" not in cleaned
        assert "Attribute VB_GlobalNameSpace" not in cleaned
        assert "Option Explicit" in cleaned
        assert "Sub Foo()" in cleaned

    def test_preserves_only_vb_name_when_keep_vb_name_true(self):
        importer = VisioVBAImporter("dummy.vsdm")
        code = (
            'Attribute VB_Name = "Mod1"\nAttribute VB_GlobalNameSpace = False\nSub Foo()\nEnd Sub\n'
        )
        cleaned = importer._strip_vba_header(code, keep_vb_name=True)
        assert 'Attribute VB_Name = "Mod1"' in cleaned
        assert "Attribute VB_GlobalNameSpace" not in cleaned

    def test_idempotent_on_already_clean_code(self):
        importer = VisioVBAImporter("dummy.vsdm")
        code = "Option Explicit\nSub Foo()\nEnd Sub\n"
        cleaned = importer._strip_vba_header(code, keep_vb_name=False)
        assert "Sub Foo()" in cleaned
        assert "Option Explicit" in cleaned

    def test_handles_form_with_nested_begin_blocks(self):
        importer = VisioVBAImporter("dummy.vsdm")
        code = (
            "VERSION 5.00\n"
            "Begin {GUID} UserForm1\n"
            '   Caption = "UserForm1"\n'
            "   Begin {GUID} Button1\n"
            '      Caption = "Click"\n'
            "   End\n"
            "End\n"
            'Attribute VB_Name = "UserForm1"\n'
            "Option Explicit\n"
            "Sub Click()\n"
            "End Sub\n"
        )
        cleaned = importer._strip_vba_header(code, keep_vb_name=False)
        assert "Begin {GUID}" not in cleaned
        assert "VERSION 5.00" not in cleaned
        assert "Attribute VB_Name" not in cleaned
        assert "Option Explicit" in cleaned
        assert "Sub Click()" in cleaned

    def test_empty_input_returns_empty(self):
        importer = VisioVBAImporter("dummy.vsdm")
        assert importer._strip_vba_header("", keep_vb_name=False).strip() == ""


# --------------------------------------------------------------------------- #
# _ensure_folder_annotation
# --------------------------------------------------------------------------- #
class TestEnsureFolderAnnotation:
    def _doc_info(self, folder_name: str = "drawing1_vsdx"):
        info = MagicMock()
        info.folder_name = folder_name
        return info

    def test_noop_when_rubberduck_disabled(self):
        importer = VisioVBAImporter("dummy.vsdm", use_rubberduck=False)
        path = Path("/ws/drawing1_vsdx/Folder/Module1.bas")
        content = "Sub Foo()\nEnd Sub\n"
        assert importer._ensure_folder_annotation(content, path, self._doc_info()) == content

    def test_injects_annotation_for_nested_path(self):
        importer = VisioVBAImporter("dummy.vsdm", use_rubberduck=True)
        path = Path("/ws/drawing1_vsdx/Foo/Bar/Module1.bas")
        content = 'Attribute VB_Name = "Module1"\nOption Explicit\n'
        new = importer._ensure_folder_annotation(content, path, self._doc_info())
        assert '\'@Folder("Foo.Bar")' in new
        assert "Option Explicit" in new

    def test_replaces_stale_annotation(self):
        importer = VisioVBAImporter("dummy.vsdm", use_rubberduck=True)
        path = Path("/ws/drawing1_vsdx/New/Module1.bas")
        content = 'Attribute VB_Name = "Module1"\n\'@Folder("Old.Path")\nOption Explicit\n'
        new = importer._ensure_folder_annotation(content, path, self._doc_info())
        assert '\'@Folder("New")' in new
        assert "Old.Path" not in new

    def test_no_annotation_at_document_root(self):
        importer = VisioVBAImporter("dummy.vsdm", use_rubberduck=True)
        path = Path("/ws/drawing1_vsdx/Module1.bas")
        content = "Option Explicit\n"
        new = importer._ensure_folder_annotation(content, path, self._doc_info())
        assert "@Folder" not in new


# --------------------------------------------------------------------------- #
# _find_document_for_file
# --------------------------------------------------------------------------- #
class TestFindDocumentForFile:
    def test_direct_parent_match_outside_rubberduck_mode(self):
        importer = VisioVBAImporter("dummy.vsdm", use_rubberduck=False)
        marker = MagicMock(name="doc")
        importer.document_map = {"drawing1": marker}
        path = Path("/ws/drawing1/Module1.bas")
        assert importer._find_document_for_file(path) is marker

    def test_rubberduck_walks_up_to_find_document(self):
        importer = VisioVBAImporter("dummy.vsdm", use_rubberduck=True)
        marker = MagicMock(name="doc")
        importer.document_map = {"drawing1": marker}
        path = Path("/ws/drawing1/Sub/Sub2/Module1.bas")
        assert importer._find_document_for_file(path) is marker

    def test_rubberduck_max_walk_depth_avoids_runaway_traversal(self):
        importer = VisioVBAImporter("dummy.vsdm", use_rubberduck=True)
        importer.document_map = {"drawing1": MagicMock()}
        # Build a deeply nested path with no match at any level.
        deep = Path("/" + "/".join(f"d{i}" for i in range(20)) + "/Module1.bas")
        # Should not raise, should not match.
        assert importer._find_document_for_file(deep) is None

    def test_non_rubberduck_falls_back_to_main_document(self):
        importer = VisioVBAImporter("dummy.vsdm", use_rubberduck=False)
        importer.document_map = {"drawing1": MagicMock()}
        importer.doc_manager = MagicMock()
        main_doc = MagicMock(name="main")
        importer.doc_manager.get_main_document.return_value = main_doc
        path = Path("/ws/unrelated/Module1.bas")
        assert importer._find_document_for_file(path) is main_doc

    def test_rubberduck_returns_none_when_no_match(self):
        importer = VisioVBAImporter("dummy.vsdm", use_rubberduck=True)
        importer.document_map = {"drawing1": MagicMock()}
        path = Path("/ws/elsewhere/Module1.bas")
        assert importer._find_document_for_file(path) is None
