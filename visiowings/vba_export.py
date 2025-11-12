"""VBA Module Export functionality: improved header stripping and hash-based change detection
Now supports multiple documents (drawings + stencils)
Enhanced: Remove local module files if deleted in Visio
"""
import win32com.client
import os
from pathlib import Path
import re
import hashlib
from .document_manager import VisioDocumentManager, VisioDocumentInfo

class VisioVBAExporter:
    def __init__(self, visio_file_path, debug=False):
        self.visio_file_path = visio_file_path
        self.visio_app = None
        self.doc = None
        self.debug = debug
        self.doc_manager = None
    
    def connect_to_visio(self, silent=False):
        try:
            self.doc_manager = VisioDocumentManager(self.visio_file_path, debug=self.debug)
            if not self.doc_manager.connect_to_visio():
                return False
            self.visio_app = self.doc_manager.visio_app
            self.doc = self.doc_manager.main_doc
            if not silent:
                self.doc_manager.print_summary()
            return True
        except Exception as e:
            if not silent:
                print(f"❌ Error connecting to Visio: {e}")
            return False
    def _strip_vba_header_file(self, file_path):
        try:
            text = Path(file_path).read_text(encoding="utf-8")
            header_markers = [
                'Option Explicit', 'Option Compare', 'Option Base', 'Sub ',
                'Function ', 'Property ', 'Public ', 'Private ', 'Dim '
            ]
            lines = text.splitlines()
            code_start = 0
            for i, line in enumerate(lines):
                stripped = line.strip()
                if not stripped or stripped.startswith(('VERSION', 'Begin', 'End', 'Attribute ', "'")):
                    continue
                if any(marker in line for marker in header_markers):
                    code_start = i
                    break
            new_text = '\n'.join(lines[code_start:])
            Path(file_path).write_text(new_text, encoding="utf-8")
            if self.debug:
                removed_lines = code_start
                if removed_lines > 0:
                    print(f"[DEBUG] {removed_lines} header lines removed from {file_path.name}")
        except Exception as e:
            if self.debug:
                print(f"[DEBUG] Error during header stripping: {e}")
            pass
    def _module_content_hash(self, vb_project):
        try:
            code_parts = []
            for comp in vb_project.VBComponents:
                cm = comp.CodeModule
                if cm.CountOfLines > 0:
                    code = cm.Lines(1, cm.CountOfLines)
                    code_parts.append(f"{comp.Name}:{code}")
            hash_input = ''.join(code_parts)
            content_hash = hashlib.md5(hash_input.encode()).hexdigest()
            if self.debug:
                print(f"[DEBUG] Hash calculated: {content_hash[:8]}... ({len(code_parts)} modules)")
            return content_hash
        except Exception as e:
            if self.debug:
                print(f"[DEBUG] Error during hash calculation: {e}")
            return None
    def _export_document_modules(self, doc_info, output_dir, last_hash=None):
        try:
            vb_project = doc_info.doc.VBProject
            current_hash = self._module_content_hash(vb_project)
            if last_hash and last_hash == current_hash:
                if self.debug:
                    print(f"[DEBUG] {doc_info.name}: Hashes identical - no export")
                # Even if no export, check if Visio modules were deleted!
                self._sync_deleted_modules(doc_info, output_dir, vb_project)
                return [], current_hash
            doc_output_path = Path(output_dir) / doc_info.folder_name
            doc_output_path.mkdir(parents=True, exist_ok=True)
            exported_files = []
            visio_module_names = set()
            ext_map = {1: '.bas', 2: '.cls', 3: '.frm', 100: '.cls'}
            for component in vb_project.VBComponents:
                ext = ext_map.get(component.Type, '.bas')
                file_name = f"{component.Name}{ext}"
                file_path = doc_output_path / file_name
                component.Export(str(file_path))
                if component.Type in [1, 2, 100]:
                    self._strip_vba_header_file(file_path)
                exported_files.append(file_path)
                visio_module_names.add(component.Name.lower())
                print(f"✓ Exported: {doc_info.folder_name}/{file_name}")
            # After export, sync deleted local files
            self._sync_deleted_modules(doc_info, output_dir, vb_project, visio_module_names)
            return exported_files, current_hash
        except Exception as e:
            print(f"❌ Error exporting {doc_info.name}: {e}")
            if self.debug:
                import traceback
                traceback.print_exc()
            return [], None
    def _sync_deleted_modules(self, doc_info, output_dir, vb_project, visio_module_names=None):
        doc_output_path = Path(output_dir) / doc_info.folder_name
        local_files = list(doc_output_path.glob("*.bas")) + list(doc_output_path.glob("*.cls")) + list(doc_output_path.glob("*.frm"))
        if visio_module_names is None:
            visio_module_names = set(comp.Name.lower() for comp in vb_project.VBComponents)
        for file in local_files:
            filename = file.stem.lower()
            if filename not in visio_module_names:
                try:
                    file.unlink()
                    print(f"✓ Removed local file: {doc_info.folder_name}/{file.name} (missing in Visio)")
                except Exception as e:
                    print(f"⚠️  Could not remove local file: {file} ({e})")
    def export_modules(self, output_dir, last_hashes=None):
        if not self.doc_manager:
            print("❌ No document manager initialized")
            return {}, {}
        if last_hashes is None:
            last_hashes = {}
        all_exported = {}
        all_hashes = {}
        documents = self.doc_manager.get_all_documents_with_vba()
        if not documents:
            print("⚠️  No documents with VBA code found")
            return {}, {}
        try:
            output_path = Path(output_dir)
            output_path.mkdir(exist_ok=True)
            for doc_info in documents:
                if self.debug:
                    print(f"[DEBUG] Exporting {doc_info.name}...")
                last_hash = last_hashes.get(doc_info.folder_name)
                exported_files, current_hash = self._export_document_modules(
                    doc_info, 
                    output_dir, 
                    last_hash
                )
                if exported_files or current_hash:
                    all_exported[doc_info.folder_name] = exported_files
                    all_hashes[doc_info.folder_name] = current_hash
            return all_exported, all_hashes
        except Exception as e:
            print(f"❌ Error during export: {e}")
            if self.debug:
                import traceback
                traceback.print_exc()
            else:
                print("")
                print("⚠️  Make sure the following setting is enabled in Visio:")
                print("   File → Options → Trust Center → Trust Center Settings")
                print("   → Macro Settings → 'Trust access to the VBA project object model'")
            return {}, {}
