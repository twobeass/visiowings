"""VBA Module Export functionality: improved header stripping and hash-based change detection
Now supports multiple documents (drawings + stencils)
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
        """Connect to Visio application and discover all documents
        
        Args:
            silent: If True, suppress connection success messages (used in polling)
        """
        try:
            # Use DocumentManager for multi-document support
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
                print(f"❌ Fehler beim Verbinden mit Visio: {e}")
            return False
    
    def _strip_vba_header_file(self, file_path):
        """Remove VBA header from exported file"""
        try:
            text = Path(file_path).read_text(encoding="utf-8")
            
            # More comprehensive header markers
            header_markers = [
                'Option Explicit',
                'Option Compare',
                'Option Base',
                'Sub ',
                'Function ',
                'Property ',
                'Public ',
                'Private ',
                'Dim '
            ]
            
            lines = text.splitlines()
            code_start = 0
            
            # Find first line that contains actual code
            for i, line in enumerate(lines):
                stripped = line.strip()
                # Skip empty lines and VBA metadata
                if not stripped or stripped.startswith(('VERSION', 'Begin', 'End', 'Attribute ', "'")):
                    continue
                # Check if this is actual code
                if any(marker in line for marker in header_markers):
                    code_start = i
                    break
            
            new_text = '\n'.join(lines[code_start:])
            Path(file_path).write_text(new_text, encoding="utf-8")
            
            if self.debug:
                removed_lines = code_start
                if removed_lines > 0:
                    print(f"[DEBUG] {removed_lines} Header-Zeilen entfernt aus {file_path.name}")
        
        except Exception as e:
            if self.debug:
                print(f"[DEBUG] Fehler beim Header-Stripping: {e}")
            pass
    
    def _module_content_hash(self, vb_project):
        """Generate content hash of all modules for change detection"""
        try:
            code_parts = []
            for comp in vb_project.VBComponents:
                cm = comp.CodeModule
                # Only hash actual code, not headers
                if cm.CountOfLines > 0:
                    code = cm.Lines(1, cm.CountOfLines)
                    code_parts.append(f"{comp.Name}:{code}")
            
            hash_input = ''.join(code_parts)
            content_hash = hashlib.md5(hash_input.encode()).hexdigest()
            
            if self.debug:
                print(f"[DEBUG] Hash berechnet: {content_hash[:8]}... ({len(code_parts)} Module)")
            
            return content_hash
        except Exception as e:
            if self.debug:
                print(f"[DEBUG] Fehler bei Hash-Berechnung: {e}")
            return None
    
    def _export_document_modules(self, doc_info, output_dir, last_hash=None):
        """Export VBA modules from a single document
        
        Args:
            doc_info: VisioDocumentInfo instance
            output_dir: Base output directory (will create subdirectory for document)
            last_hash: Last known hash for this document
        
        Returns:
            tuple: (list of exported files, current hash)
        """
        try:
            vb_project = doc_info.doc.VBProject
            
            # Calculate current hash
            current_hash = self._module_content_hash(vb_project)
            
            # Check if content actually changed
            if last_hash and last_hash == current_hash:
                if self.debug:
                    print(f"[DEBUG] {doc_info.name}: Hashes identisch - kein Export")
                return [], current_hash
            
            # Create subdirectory for this document
            doc_output_path = Path(output_dir) / doc_info.folder_name
            doc_output_path.mkdir(parents=True, exist_ok=True)
            
            exported_files = []
            
            for component in vb_project.VBComponents:
                # Map component types to file extensions
                ext_map = {
                    1: '.bas',    # vbext_ct_StdModule
                    2: '.cls',    # vbext_ct_ClassModule
                    3: '.frm',    # vbext_ct_MSForm
                    100: '.cls'   # vbext_ct_Document
                }
                
                ext = ext_map.get(component.Type, '.bas')
                file_name = f"{component.Name}{ext}"
                file_path = doc_output_path / file_name
                
                # Export module
                component.Export(str(file_path))
                
                # Remove VBA header for standard modules and class modules
                if component.Type in [1, 2, 100]:
                    self._strip_vba_header_file(file_path)
                
                exported_files.append(file_path)
                print(f"✓ Exportiert: {doc_info.folder_name}/{file_name}")
            
            return exported_files, current_hash
        
        except Exception as e:
            print(f"❌ Fehler beim Exportieren von {doc_info.name}: {e}")
            if self.debug:
                import traceback
                traceback.print_exc()
            return [], None
    
    def export_modules(self, output_dir, last_hashes=None):
        """Export VBA modules from all documents with VBA code
        
        Args:
            output_dir: Base output directory
            last_hashes: Dict mapping document folder names to their last hash
        
        Returns:
            tuple: (dict of {doc_folder: exported_files}, dict of {doc_folder: hash})
        """
        if not self.doc_manager:
            print("❌ Kein Dokument-Manager initialisiert")
            return {}, {}
        
        if last_hashes is None:
            last_hashes = {}
        
        all_exported = {}
        all_hashes = {}
        
        documents = self.doc_manager.get_all_documents_with_vba()
        
        if not documents:
            print("⚠️  Keine Dokumente mit VBA-Code gefunden")
            return {}, {}
        
        try:
            output_path = Path(output_dir)
            output_path.mkdir(exist_ok=True)
            
            for doc_info in documents:
                if self.debug:
                    print(f"[DEBUG] Exportiere {doc_info.name}...")
                
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
            print(f"❌ Fehler beim Exportieren: {e}")
            if self.debug:
                import traceback
                traceback.print_exc()
            else:
                print("")
                print("⚠️  Stelle sicher, dass in Visio folgende Einstellung aktiviert ist:")
                print("   Datei → Optionen → Trust Center → Trust Center-Einstellungen")
                print("   → Makroeinstellungen → 'Zugriff auf VBA-Projektobjektmodell vertrauen'")
            
            return {}, {}
