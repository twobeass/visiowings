import pytest
from pathlib import Path
from unittest.mock import MagicMock
from visiowings.vba_export import VisioVBAExporter
from visiowings.vba_import import VisioVBAImporter

class TestRubberduckIntegration:

    def test_extract_folder_annotation(self):
        exporter = VisioVBAExporter("dummy.vsdm", use_rubberduck=True)

        # Case 1: Standard annotation (comment)
        code = """Attribute VB_Name = "Module1"
'@Folder("Main.Sub.Deep")
Option Explicit
Sub Test()
End Sub"""
        path = exporter._extract_folder_annotation(code)
        assert path == "Main/Sub/Deep"

        # Case 2: No annotation
        code_none = """Attribute VB_Name = "Module1"
Option Explicit"""
        path = exporter._extract_folder_annotation(code_none)
        assert path is None

        # Case 3: Spaces and variants (comment)
        code_space = """'@Folder ( "My Folder" )"""
        path = exporter._extract_folder_annotation(code_space)
        assert path == "My Folder"

    def test_ensure_folder_annotation_injection(self):
        importer = VisioVBAImporter("dummy.vsdm", use_rubberduck=True)
        doc_info = MagicMock()
        doc_info.folder_name = "Drawing1_vsdx"

        # Case 1: New file deep in structure
        # Path: /workspace/Drawing1_vsdx/Folder/Sub/Module1.bas
        file_path = Path("/workspace/Drawing1_vsdx/Folder/Sub/Module1.bas")
        content = """Attribute VB_Name = "Module1"
Option Explicit
Sub Test()
End Sub"""

        new_content = importer._ensure_folder_annotation(content, file_path, doc_info)
        # Expect comment prefix
        assert "'@Folder(\"Folder.Sub\")" in new_content
        assert 'Option Explicit' in new_content

        # Check position: Should be injected
        assert "'@Folder(\"Folder.Sub\")" in new_content

    def test_ensure_folder_annotation_update(self):
        importer = VisioVBAImporter("dummy.vsdm", use_rubberduck=True)
        doc_info = MagicMock()
        doc_info.folder_name = "Drawing1_vsdx"

        # Case 2: Existing incorrect annotation (comment)
        file_path = Path("/workspace/Drawing1_vsdx/NewLocation/Module1.bas")
        content = """Attribute VB_Name = "Module1"
'@Folder("OldLocation")
Option Explicit"""

        new_content = importer._ensure_folder_annotation(content, file_path, doc_info)
        assert "'@Folder(\"NewLocation\")" in new_content
        assert "OldLocation" not in new_content

    def test_ensure_folder_annotation_root(self):
        importer = VisioVBAImporter("dummy.vsdm", use_rubberduck=True)
        doc_info = MagicMock()
        doc_info.folder_name = "Drawing1_vsdx"

        # Case 3: File at root of document folder
        file_path = Path("/workspace/Drawing1_vsdx/Module1.bas")
        content = "Option Explicit"

        new_content = importer._ensure_folder_annotation(content, file_path, doc_info)
        assert '@Folder' not in new_content

    def test_find_document_for_file_recursive(self):
        importer = VisioVBAImporter("dummy.vsdm", use_rubberduck=True)

        # Setup mock document map
        mock_doc = MagicMock()
        importer.document_map = {"Drawing1_vsdx": mock_doc}

        # Test deep path
        deep_path = Path("/workspace/Drawing1_vsdx/Folder/Sub/Module.bas")
        found_doc = importer._find_document_for_file(deep_path)
        assert found_doc == mock_doc

        # Test root path
        root_path = Path("/workspace/Drawing1_vsdx/Module.bas")
        found_doc = importer._find_document_for_file(root_path)
        assert found_doc == mock_doc

        # Test unrelated path (should fall back to main doc mock)
        importer.doc_manager = MagicMock()
        main_doc_mock = MagicMock()
        importer.doc_manager.get_main_document.return_value = main_doc_mock

        other_path = Path("/workspace/Other/Module.bas")
        found_doc = importer._find_document_for_file(other_path)
        assert found_doc == main_doc_mock
