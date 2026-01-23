"""Document Manager for handling multiple Visio documents and stencils

Supports:
- Main drawing files (.vsdx, .vsdm)
- Stencil files (.vssm, .vssx)
- Template files (.vstm, .vstx)

Auto-detects all open documents and organizes VBA code by document.
"""
import re
from pathlib import Path

import pythoncom
import win32com.client


class VisioDocumentType:
    """Visio document type constants"""
    DRAWING = 1    # visTypeDrawing
    STENCIL = 2    # visTypeStencil
    TEMPLATE = 3   # visTypeTemplate
    
def sanitize_document_name(name):
    """Sanitize document name for use as folder name"""
    # Use full name first to avoid path splitting issues on special chars
    # Remove extension manually or use Path if safe
    # On Windows paths, special chars like | or ? are invalid in filenames anyway,
    # but if we get a raw string, we should handle it robustly.

    # Just strip extension first
    if '.' in name:
        name = name.rsplit('.', 1)[0]

    # Replace invalid characters with underscore
    name = re.sub(r'[<>:"/\\|?*]', '_', name)

    # Remove spaces
    name = name.replace(' ', '_').lower()

    # Remove consecutive underscores
    name = re.sub(r'_+', '_', name)

    # Remove leading/trailing underscores
    name = name.strip('_')

    return name or 'document'

class VisioDocumentInfo:
    """Information about a Visio document"""

    def __init__(self, doc, debug=False):
        self.doc = doc
        self.debug = debug
        self.name = doc.Name
        self.full_name = doc.FullName
        self.type = doc.Type
        self.has_vba = self._check_has_vba()
        self.folder_name = sanitize_document_name(self.name)

    def _check_has_vba(self):
        """Check if document has VBA code"""
        try:
            vb_project = self.doc.VBProject
            if vb_project and vb_project.VBComponents.Count > 0:
                return True
        except:
            pass
        return False

    def get_type_name(self):
        """Get human-readable document type name"""
        type_names = {
            VisioDocumentType.DRAWING: 'Drawing',
            VisioDocumentType.STENCIL: 'Stencil',
            VisioDocumentType.TEMPLATE: 'Template'
        }
        return type_names.get(self.type, 'Unknown')

    def __repr__(self):
        return f"VisioDocumentInfo(name='{self.name}', type={self.get_type_name()}, has_vba={self.has_vba})"

class VisioDocumentManager:
    """Manages multiple Visio documents and stencils"""

    def __init__(self, main_file_path, debug=False):
        self.main_file_path = Path(main_file_path)
        self.debug = debug
        self.visio_app = None
        self.main_doc = None
        self.documents = []  # List of VisioDocumentInfo

    def connect_to_visio(self):
        """Connect to Visio and discover all open documents"""
        try:
            # Ensure COM is initialized in this thread
            try:
                pythoncom.CoInitialize()
                if self.debug:
                    print("[DEBUG] COM initialized in document_manager")
            except:
                # Already initialized, that's fine
                if self.debug:
                    print("[DEBUG] COM already initialized in document_manager")
                pass

            self.visio_app = win32com.client.Dispatch("Visio.Application")

            # Find main document
            main_doc_found = False
            main_file_name = self.main_file_path.name.lower()

            # Debug: Show all open documents
            if self.debug:
                print(f"[DEBUG] Looking for: {str(self.main_file_path)}")
                print(f"[DEBUG] Filename: {main_file_name}")
                print("[DEBUG] Open documents in Visio:")

            for doc in self.visio_app.Documents:
                doc_full_path = doc.FullName.lower()
                doc_name = Path(doc_full_path).name.lower()

                if self.debug:
                    print(f"[DEBUG]   - {doc.Name}")
                    print(f"[DEBUG]     Path: {doc_full_path}")

                # Strategy 1: Exact path match
                if doc_full_path == str(self.main_file_path).lower():
                    self.main_doc = doc
                    main_doc_found = True
                    if self.debug:
                        print("[DEBUG]     ✓ MATCHED (exact path)")
                    break

                # Strategy 2: Filename match (for OneDrive, SharePoint, etc.)
                if doc_name == main_file_name:
                    self.main_doc = doc
                    main_doc_found = True
                    if self.debug:
                        print("[DEBUG]     ✓ MATCHED (by filename)")
                        print("[DEBUG]     Note: Using filename match because paths differ")
                        print(f"[DEBUG]     Expected: {str(self.main_file_path)}")
                        print(f"[DEBUG]     Actual:   {doc_full_path}")
                    break

            if not main_doc_found:
                print(f"⚠️  Document not found: {self.main_file_path.name}")
                print("\n   Open documents in Visio:")
                for doc in self.visio_app.Documents:
                    print(f"     - {doc.Name}")
                print("\n   Tip: Make sure the document is open in Visio")
                return False

            # Discover all open documents
            self._discover_documents()

            if self.debug:
                print(f"[DEBUG] Documents found: {len(self.documents)}")
                for doc_info in self.documents:
                    print(f"[DEBUG]   - {doc_info}")

            return True

        except Exception as e:
            print(f"❌ Error connecting to Visio: {e}")
            if self.debug:
                import traceback
                traceback.print_exc()
            return False

    def _discover_documents(self):
        """Discover all open documents with VBA code"""
        self.documents = []

        try:
            for doc in self.visio_app.Documents:
                doc_info = VisioDocumentInfo(doc, debug=self.debug)

                # Only include documents with VBA code
                if doc_info.has_vba:
                    self.documents.append(doc_info)

                    if self.debug:
                        print(f"[DEBUG] VBA found in: {doc_info.name} ({doc_info.get_type_name()})")

        except Exception as e:
            if self.debug:
                print(f"[DEBUG] Error during document discovery: {e}")

    def get_main_document(self):
        """Get main document info"""
        for doc_info in self.documents:
            if doc_info.doc == self.main_doc:
                return doc_info
        return None

    def get_stencils(self):
        """Get all stencil documents"""
        return [d for d in self.documents if d.type == VisioDocumentType.STENCIL]

    def get_all_documents_with_vba(self):
        """Get all documents that contain VBA code"""
        return self.documents

    def is_multi_document(self):
        """Check if multiple documents with VBA are open"""
        return len(self.documents) > 1

    def print_summary(self):
        """Print summary of discovered documents"""
        if not self.documents:
            print("⚠️  No documents with VBA code found")
            return

        print(f"\n📚 Documents with VBA found: {len(self.documents)}")

        main_doc_info = self.get_main_document()
        if main_doc_info:
            print(f"   📄 Main document: {main_doc_info.name} ({main_doc_info.get_type_name()})")

        stencils = self.get_stencils()
        if stencils:
            print(f"   📋 Stencils: {len(stencils)}")
            for stencil in stencils:
                print(f"      - {stencil.name}")

        print()
