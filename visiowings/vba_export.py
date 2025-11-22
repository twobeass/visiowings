"""VBA Module Export functionality: improved header stripping and hash-based change detection
Now supports multiple documents (drawings + stencils)
Enhanced: Remove local module files if deleted in Visio
Fixed: Proper file comparison with normalization to avoid false positives
"""
import win32com.client
import os
from pathlib import Path
import re
import hashlib
import tempfile
from .document_manager import VisioDocumentManager, VisioDocumentInfo
import difflib

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
    
    def _normalize_content(self, content):
        """Normalize content for comparison by removing insignificant differences"""
        # Split into lines
        lines = content.splitlines()
        
        # Strip trailing whitespace from each line and remove empty lines at start/end
        normalized_lines = [line.rstrip() for line in lines]
        
        # Remove leading empty lines
        while normalized_lines and not normalized_lines[0]:
            normalized_lines.pop(0)
        
        # Remove trailing empty lines
        while normalized_lines and not normalized_lines[-1]:
            normalized_lines.pop()
        
        # Join with consistent line ending
        return '\n'.join(normalized_lines)
    
    def _strip_vba_header_export(self, text, keep_vb_name=True):
        import re
        lines = text.splitlines()
        filtered_lines = []
        for line in lines:
            s = line.strip()
            if s.startswith('VERSION') or s.startswith('MultiUse'):
                continue
            if re.match(r'^Begin( |\{)', s):
                continue
            if s == 'End':
                continue
            if s.startswith('Attribute '):
                if keep_vb_name and 'VB_Name' in line:
                    filtered_lines.append(line)
                continue
            filtered_lines.append(line)
        return '\n'.join(filtered_lines)

    def _strip_and_convert(self, file_path):
        # Read cp1252 (Visio export), clean headers, warn if transcoding loses data
        try:
            raw = Path(file_path).read_text(encoding="cp1252")
            cleaned = self._strip_vba_header_export(raw, keep_vb_name=True)
            try:
                cleaned.encode('utf-8')  # if this errors, warn below
            except UnicodeEncodeError as e:
                print(f"⚠️  Warning: {file_path.name} contains characters not representable in utf-8: {e}")
            Path(file_path).write_text(cleaned, encoding="utf-8")
            return cleaned
        except Exception as e:
            print(f"⚠️  Encoding or cleaning error for {file_path}: {e}")
            return None
        

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
    
    def _compare_module_content(self, local_path, component):
        """Compare local file with Visio module content using normalization
        Returns: (are_different, local_hash, visio_hash)
        """
        try:
            # Read local file
            local_content = local_path.read_text(encoding="utf-8")
            
            # Get Visio module content
            cm = component.CodeModule
            if cm.CountOfLines > 0:
                visio_content = cm.Lines(1, cm.CountOfLines)
            else:
                visio_content = ""
            
            # Normalize both for comparison
            local_normalized = self._normalize_content(local_content)
            visio_normalized = self._normalize_content(visio_content)
            
            # Calculate hashes for debugging
            local_hash = hashlib.md5(local_normalized.encode()).hexdigest()[:8]
            visio_hash = hashlib.md5(visio_normalized.encode()).hexdigest()[:8]
            
            are_different = local_normalized != visio_normalized
            
            if self.debug and are_different:
                print(f"[DEBUG] Content differs: {local_path.name}")
                print(f"[DEBUG]   Local hash:  {local_hash}")
                print(f"[DEBUG]   Visio hash:  {visio_hash}")
            
            return are_different, local_hash, visio_hash
            
        except Exception as e:
            if self.debug:
                print(f"[DEBUG] Error comparing {local_path.name}: {e}")
            # On error, assume different to be safe
            return True, None, None
    
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
            
            # Check for files with local changes
            files_with_changes = {}
            ext_map = {1: '.bas', 2: '.cls', 3: '.frm', 100: '.cls'}
            for file_name, file_path, visio_clean, local_clean in files_with_changes:
                print(f"\n⚠️  File differs: {file_name}")
                diff_lines = list(
                    difflib.unified_diff(
                        local_clean.splitlines(),
                        visio_clean.splitlines(),
                        fromfile='Local',
                        tofile='Visio',
                        lineterm=''
                    )
                )
                if diff_lines:
                    print('\n'.join(diff_lines))
                response = input(f"Overwrite local file '{file_name}' with Visio content? (y/N): ").strip().lower()
                if response not in ('y', 'yes'):
                    print(f"⊘ Skipped: {file_name}")
                    continue

            for component in vb_project.VBComponents:
                ext = ext_map.get(component.Type, '.bas')
                file_name = f"{component.Name}{ext}"
                file_path = doc_output_path / file_name
                
                # If file exists locally and is a code module, check for changes
                if file_path.exists() and component.Type in [1, 2, 100]:
                    are_different, local_hash, visio_hash = self._compare_module_content(
                        file_path, component
                    )
                    
                    if are_different:
                        files_with_changes[file_name] = {
                            'path': file_path,
                            'component': component,
                            'local_hash': local_hash,
                            'visio_hash': visio_hash
                        }
            
            # If there are files with local changes, handle them interactively
            files_to_skip = set()
            if files_with_changes:
                print(f"\n⚠️  Local changes detected in {doc_info.name}:")
                for fname in files_with_changes.keys():
                    print(f"   - {doc_info.folder_name}/{fname}")
                
                print(f"\nOptions:")
                print(f"  o - Overwrite all with Visio content")
                print(f"  s - Skip changed files (keep local changes)")
                print(f"  i - Interactive (choose per file)")
                print(f"  c - Cancel export for this document")
                
                response = input(f"\nChoose action (o/s/i/C): ").strip().lower()
                
                if response == 'o':
                    # Overwrite all - proceed normally
                    print(f"✓ Will overwrite {len(files_with_changes)} file(s)")
                elif response == 's':
                    # Skip all changed files
                    files_to_skip = set(files_with_changes.keys())
                    print(f"✓ Will skip {len(files_to_skip)} changed file(s)")
                elif response == 'i':
                    # Interactive mode
                    for fname, info in files_with_changes.items():
                        print(f"\n{doc_info.folder_name}/{fname}")
                        if self.debug:
                            print(f"  Local:  {info['local_hash']}")
                            print(f"  Visio:  {info['visio_hash']}")
                        choice = input(f"  Overwrite? (y/N): ").strip().lower()
                        if choice not in ('y', 'yes'):
                            files_to_skip.add(fname)
                else:
                    # Cancel (default)
                    print(f"❌ Export cancelled for {doc_info.name}")
                    return [], None
            
            # Proceed with export
            exported_files = []
            visio_module_names = set()
            skipped_count = 0
            
            for component in vb_project.VBComponents:
                ext = ext_map.get(component.Type, '.bas')
                file_name = f"{component.Name}{ext}"
                file_path = doc_output_path / file_name
                
                # Skip if user chose to keep local changes
                if file_name in files_to_skip:
                    if self.debug:
                        print(f"⊘ Skipped: {doc_info.folder_name}/{file_name} (local changes preserved)")
                    skipped_count += 1
                    visio_module_names.add(component.Name.lower())
                    continue
                
                # Export the module
                component.Export(str(file_path))
                if component.Type in [1, 2, 100]:
                    self._strip_and_convert(file_path)
                exported_files.append(file_path)
                visio_module_names.add(component.Name.lower())
                print(f"✓ Exported: {doc_info.folder_name}/{file_name}")
            
            if skipped_count > 0:
                print(f"ℹ️  Skipped {skipped_count} file(s) with local changes")
            
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
        
        # Collect files to delete
        files_to_delete = []
        for file in local_files:
            filename = file.stem.lower()
            if filename not in visio_module_names:
                files_to_delete.append(file)
        
        # If there are files to delete, ask user
        if files_to_delete:
            print(f"\n⚠️  The following local files are missing in Visio:")
            for file in files_to_delete:
                print(f"   - {doc_info.folder_name}/{file.name}")
            
            print(f"\nOptions:")
            print(f"  d - Delete local files")
            print(f"  i - Import to Visio")
            print(f"  k - Keep local files (default)")
            response = input(f"\nChoose action (d/i/K): ").strip().lower()
            
            if response == 'd':
                for file in files_to_delete:
                    try:
                        file.unlink()
                        print(f"✓ Removed local file: {doc_info.folder_name}/{file.name}")
                    except Exception as e:
                        print(f"⚠️  Could not remove local file: {file} ({e})")
            elif response == 'i':
                print(f"\n📤 Importing {len(files_to_delete)} file(s) to Visio...")
                for file in files_to_delete:
                    try:
                        vb_project.VBComponents.Import(str(file))
                        print(f"✓ Imported to Visio: {doc_info.folder_name}/{file.name}")
                    except Exception as e:
                        print(f"✗ Error importing {file.name}: {e}")
            else:
                print(f"ℹ️  Kept {len(files_to_delete)} local file(s)")
    
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
