"""
VBA Module Export functionality: improved Unicode-compliant export.
Now extracts VBA code modules directly as Unicode, preventing character corruption.
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
        lines = content.splitlines()
        normalized_lines = [line.rstrip() for line in lines]
        while normalized_lines and not normalized_lines[0]:
            normalized_lines.pop(0)
        while normalized_lines and not normalized_lines[-1]:
            normalized_lines.pop()
        return '\n'.join(normalized_lines)

    def _strip_vba_header_export(self, text, keep_vb_name=True):
        lines = text.splitlines()
        filtered_lines = []
        in_begin_block = False
        for line in lines:
            s = line.strip()
            if s.startswith('VERSION'):
                continue
            if re.match(r'^BEGIN(\s|$)', s, re.IGNORECASE):
                in_begin_block = True
                continue
            if in_begin_block and re.match(r'^END$', s, re.IGNORECASE):
                in_begin_block = False
                continue
            if in_begin_block:
                continue
            if s.startswith('MultiUse'):
                continue
            if s.startswith('Attribute '):
                if keep_vb_name and 'VB_Name' in line:
                    filtered_lines.append(line)
                continue
            filtered_lines.append(line)
        return '\n'.join(filtered_lines)

    def _strip_and_convert(self, file_path):
        # Maintains legacy support (used only for forms or pure exports)
        try:
            raw = Path(file_path).read_text(encoding="cp1252")
            cleaned = self._strip_vba_header_export(raw, keep_vb_name=True)
            try:
                cleaned.encode('utf-8')
            except UnicodeEncodeError as e:
                print(f"⚠️  Warning: {file_path.name} contains characters not representable in utf-8: {e}")
            Path(file_path).write_text(cleaned, encoding="utf-8")
            return cleaned
        except Exception as e:
            print(f"⚠️  Encoding or cleaning error for {file_path}: {e}")
            return None

    def _read_local_file_with_fallback(self, file_path):
        try:
            return file_path.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            if self.debug:
                print(f"[DEBUG] {file_path.name} not UTF-8, reading as cp1252 and converting")
            try:
                content = file_path.read_text(encoding="cp1252")
                file_path.write_text(content, encoding="utf-8")
                if self.debug:
                    print(f"[DEBUG] Converted {file_path.name} to UTF-8")
                return content
            except Exception as fallback_error:
                raise Exception(f"Cannot read {file_path.name} as UTF-8 or cp1252: {fallback_error}")

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
        try:
            local_content = self._read_local_file_with_fallback(local_path)
            local_normalized = self._strip_vba_header_export(local_content, keep_vb_name=False)
            cm = component.CodeModule
            if cm.CountOfLines > 0:
                visio_content = cm.Lines(1, cm.CountOfLines)
            else:
                visio_content = ""
            visio_normalized = self._strip_vba_header_export(visio_content, keep_vb_name=False)
            local_final = self._normalize_content(local_normalized)
            visio_final = self._normalize_content(visio_normalized)
            local_hash = hashlib.md5(local_final.encode()).hexdigest()[:8]
            visio_hash = hashlib.md5(visio_final.encode()).hexdigest()[:8]
            are_different = local_final != visio_final
            if self.debug and are_different:
                print(f"[DEBUG] Content differs: {local_path.name}")
                print(f"[DEBUG]   Local hash:  {local_hash}")
                print(f"[DEBUG]   Visio hash:  {visio_hash}")
            return are_different, local_hash, visio_hash
        except Exception as e:
            if self.debug:
                print(f"[DEBUG] Error comparing {local_path.name}: {e}")
            return True, None, None

    def _export_document_modules(self, doc_info, output_dir, last_hash=None):
        try:
            vb_project = doc_info.doc.VBProject
            current_hash = self._module_content_hash(vb_project)
            if last_hash and last_hash == current_hash:
                if self.debug:
                    print(f"[DEBUG] {doc_info.name}: Hashes identical - no export")
                self._sync_deleted_modules(doc_info, output_dir, vb_project)
                return [], current_hash
            doc_output_path = Path(output_dir) / doc_info.folder_name
            doc_output_path.mkdir(parents=True, exist_ok=True)
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
                # KEY FIX: Export code modules directly as Unicode (never via Export!), preserves UTF-8 chars
                if component.Type in [1, 2, 100]:
                    cm = component.CodeModule
                    if cm.CountOfLines > 0:
                        code = cm.Lines(1, cm.CountOfLines)
                        cleaned = self._strip_vba_header_export(code, keep_vb_name=True)
                        Path(file_path).write_text(cleaned, encoding="utf-8")
                    else:
                        Path(file_path).write_text("", encoding="utf-8")
                elif component.Type == 3:
                    # Forms: fall back to legacy export for .frm (these are harder to rehydrate from code only)
                    component.Export(str(file_path))
                    self._strip_and_convert(file_path)
            # After export, sync deleted local files
            self._sync_deleted_modules(doc_info, output_dir, vb_project, set(comp.Name.lower() for comp in vb_project.VBComponents))
            return [Path(output_dir) / doc_info.folder_name / f"{component.Name}{ext_map.get(component.Type, '.bas')}" for component in vb_project.VBComponents], current_hash
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
        files_to_delete = []
        for file in local_files:
            filename = file.stem.lower()
            if filename not in visio_module_names:
                files_to_delete.append(file)
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
