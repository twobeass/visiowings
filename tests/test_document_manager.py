import pytest
from unittest.mock import MagicMock
from visiowings.document_manager import VisioDocumentInfo

class TestDocumentManager:
    def test_sanitize_name(self):
        # We need to mock the doc object
        mock_doc = MagicMock()
        mock_doc.Name = "Test Drawing.vsdx"
        mock_doc.FullName = "C:\\Docs\\Test Drawing.vsdx"
        mock_doc.Type = 1

        # Mock VBProject to avoid error in _check_has_vba
        mock_vb_project = MagicMock()
        mock_vb_project.VBComponents.Count = 1
        mock_doc.VBProject = mock_vb_project

        doc_info = VisioDocumentInfo(mock_doc)

        # Original logic:
        # 1. Stem: "Test Drawing"
        # 2. Lowercase: "test drawing"
        # 3. Replace space with underscore: "test_drawing"
        assert doc_info.folder_name == "test_drawing"

    def test_sanitize_name_special_chars(self):
        mock_doc = MagicMock()
        mock_doc.Name = "My/Cool|Drawing?.vsdm"
        mock_doc.Type = 1

        mock_vb_project = MagicMock()
        mock_vb_project.VBComponents.Count = 1
        mock_doc.VBProject = mock_vb_project

        doc_info = VisioDocumentInfo(mock_doc)

        # Should replace special chars with underscores and collapse them
        assert "/" not in doc_info.folder_name
        assert "|" not in doc_info.folder_name
        assert "?" not in doc_info.folder_name
        # logic: my_cool_drawing
        assert doc_info.folder_name == "my_cool_drawing"
