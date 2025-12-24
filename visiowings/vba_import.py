import os
import re
from difflib import unified_diff
from pathlib import Path

import pythoncom

from .document_manager import VisioDocumentManager
from .encoding import DEFAULT_CODEPAGE, resolve_encoding


class VisioVBAImporter:
    def __init__(self, visio_file_path, force_document=False, debug=False, silent_reconnect=False, always_yes=False, user_codepage=None):
        self.visio_file_path = visio_file_path
        self.visio_app = None
        self.doc = None
        self.force_document = force_document
        self.debug = debug
        self.silent_reconnect = silent_reconnect
        self.doc_manager = None
        self.document_map = {}
        self.always_yes = always_yes
        self.user_codepage = user_codepage
        self.codepage = DEFAULT_CODEPAGE

    def connect_to_visio(self):
        try:
            pythoncom.CoInitialize()
        except Exception as e:
            if self.debug:
                print(f"[DEBUG] COM already initialized: {type(e).__name__}: {e}")
        self.doc_manager = VisioDocumentManager(self.visio_file_path, debug=self.debug)
        if not self.doc_manager.connect_to_visio():
            return False
        self.visio_app = self.doc_manager.visio_app
        self.doc = self.doc_manager.main_doc
        if not self.doc:
            print("❌ Failed to connect to main document")
            return False

        # Resolve encoding (user-specified > document language)
        self.codepage = resolve_encoding(
            document=self.doc,
            user_codepage=self.user_codepage,
            debug=self.debug
        )

        for doc_info in self.doc_manager.get_all_documents_with_vba():
            self.document_map[doc_info.folder_name] = doc_info
        if self.debug:
            print(f"[DEBUG] Document map created: {list(self.document_map.keys())}")
        return True

    def _ensure_connection(self):
        try:
            _ = self.doc.Name
            return True
        except Exception as e:
            if self.debug and not self.silent_reconnect:
                print(f"[DEBUG] Connection lost ({type(e).__name__}: {e}), attempting to reconnect...")
            elif not self.debug and not self.silent_reconnect:
                print("🔄 Connection lost, attempting to reconnect...")
            return self.connect_to_visio()

    def _find_document_for_file(self, file_path):
        """
        Returns the Visio document object associated with this file, based on parent folder.
        If no matching document is open, returns None and warns the user.
        """
        parent_dir = file_path.parent.name
        if parent_dir in self.document_map:
            if self.debug:
                print(f"[DEBUG] File {file_path.name} belongs to document: {parent_dir}")
            return self.document_map[parent_dir]
        else:
            print(
                f"⚠️  No open Visio document found for folder '{parent_dir}'; "
                f"skipping import of '{file_path.name}' from this folder!"
            )
            return None

    def _create_temp_codepage_file(self, file_path, codepage):
        """Create a temporary file with configured encoding for VBA import."""
        import tempfile
        try:
            text = file_path.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            text = file_path.read_text(encoding=codepage)
        module_name = file_path.stem
        if "Attribute VB_Name" not in text:
            header = f'Attribute VB_Name = "{module_name}"\n'
            text = header + text
        if text and not text.endswith("\n"):
            text += "\n"
        fd, temp_path = tempfile.mkstemp(suffix=file_path.suffix, text=True)
        try:
            with os.fdopen(fd, 'w', encoding=codepage) as f:
                f.write(text)
            if self.debug:
                print(f"[DEBUG] Created temp {codepage.upper()} file: {temp_path}")
            return temp_path
        except UnicodeEncodeError:
            print(f"⚠️  Warning: {file_path.name} contains characters not supported in {codepage.upper()}")
            with os.fdopen(fd, 'w', encoding=codepage, errors='replace') as f:
                f.write(text)
            return temp_path
        except Exception:
            os.close(fd)
            os.unlink(temp_path)
            raise

    def import_module(self, file_path, edit_mode=False):
        com_initialized = False
        temp_file = None
        try:
            pythoncom.CoInitialize()
            com_initialized = True
            if self.debug:
                print("[DEBUG] COM initialized for import_module thread")
        except Exception as e:
            if self.debug:
                print(f"[DEBUG] COM already initialized in this thread: {type(e).__name__}: {e}")
        try:
            if not self.connect_to_visio():
                print("⚠️  No connection to Visio - make sure the document is open")
                return False
            file_path = Path(file_path)
            target_doc_info = self._find_document_for_file(file_path)
            if not target_doc_info:
                print(f"⚠️  No matching document found for {file_path.name}")
                return False
            vb_project = target_doc_info.doc.VBProject
            module_name = file_path.stem
            if self.debug:
                print(f"[DEBUG] Importing {file_path.name} into {target_doc_info.name}")
            component = None
            for comp in vb_project.VBComponents:
                if comp.Name == module_name:
                    component = comp
                    break
            if component and component.Type == 100:
                if self.force_document:
                    try:
                        code = file_path.read_text(encoding="utf-8")
                    except Exception:
                        code = file_path.read_text(encoding=self.codepage, errors='replace')
                    code = self._strip_vba_header(code)
                    cm = component.CodeModule
                    if cm.CountOfLines > 0:
                        cm.DeleteLines(1, cm.CountOfLines)
                    if code.strip():
                        cm.AddFromString(code)
                    print(f"✓ Imported: {target_doc_info.folder_name}/{file_path.name} (force)")
                    return True
                else:
                    print(f"⚠️  Document module '{module_name}' skipped without --force.")
                    return False
            if component:
                if not self._prompt_overwrite(module_name, file_path, component, edit_mode=edit_mode):
                    print(f"⊘ Skipped: {module_name}")
                    return False
                vb_project.VBComponents.Remove(component)
            temp_file = self._create_temp_codepage_file(file_path, self.codepage)
            vb_project.VBComponents.Import(str(temp_file))
            print(f"✓ Imported: {target_doc_info.folder_name}/{file_path.name}")
            return True
        except Exception as e:
            print(f"✗ Error importing {file_path}: {type(e).__name__}: {e}")
            if self.debug:
                import traceback
                traceback.print_exc()
            return False
        finally:
            if temp_file and temp_file != str(file_path):
                try:
                    os.unlink(temp_file)
                    if self.debug:
                        print(f"[DEBUG] Cleaned up temp file: {temp_file}")
                except Exception as e:
                    if self.debug:
                        print(f"[DEBUG] Error cleaning temp file: {type(e).__name__}: {e}")
            if com_initialized:
                try:
                    pythoncom.CoUninitialize()
                    if self.debug:
                        print("[DEBUG] COM uninitialized for import_module thread")
                except Exception as e:
                    if self.debug:
                        print(f"[DEBUG] Error uninitializing COM: {type(e).__name__}: {e}")

    def get_document_folders(self):
        """Return list of document folder names detected for import/export mapping."""
        return list(self.document_map.keys())

    def _module_type_from_ext(self, filename):
        ext = Path(filename).suffix.lower()
        if ext == ".bas":
            return "module"
        elif ext == ".cls":
            return "class"
        elif ext == ".frm":
            return "form"
        return "unknown"

    def _read_module_code(self, file_path):
        try:
            return file_path.read_text(encoding="utf-8")
        except Exception:
            try:
                return file_path.read_text(encoding=self.codepage)
            except Exception:
                return ""

    def _strip_vba_header(self, code, keep_vb_name=False):
        """Strip VBA headers with proper handling of classes and forms.

        This function removes VBA IDE metadata while preserving actual code:
        - VERSION lines
        - BEGIN...End blocks (including nested blocks in forms/classes)
        - MultiUse declarations
        - Attribute lines (except VB_Name when keep_vb_name=True)

        Args:
            code: VBA code text to process
            keep_vb_name: If True, preserves Attribute VB_Name line

        Returns:
            Cleaned VBA code with headers removed
        """
        lines = code.splitlines()
        filtered_lines = []
        begin_depth = 0  # Track nesting depth of BEGIN blocks

        # VBA code keywords that end with 'End' (case-insensitive)
        code_end_keywords = {
            'end sub', 'end function', 'end property', 'end if',
            'end with', 'end select', 'end type', 'end enum'
        }

        for line in lines:
            s = line.strip()
            s_lower = s.lower()

            # Remove VERSION lines
            if s.upper().startswith('VERSION'):
                continue

            # Detect BEGIN block start (with any parameters)
            # Matches: BEGIN, BEGIN VB.Form, BEGIN {GUID} ControlName, etc.
            if re.match(r'^BEGIN\s+', s, re.IGNORECASE) or s_lower == 'begin':
                begin_depth += 1
                if self.debug:
                    print(f"[DEBUG] BEGIN detected (depth={begin_depth}): {s[:50]}")
                continue

            # Detect END of BEGIN block
            # Must be standalone 'End' or 'End Begin', not 'End Sub', 'End Function', etc.
            if begin_depth > 0:
                # Check if this is a block terminator END (not a code keyword)
                is_block_end = (
                    s_lower == 'end' or
                    s_lower == 'end begin' or
                    (s_lower.startswith('end ') and s_lower not in code_end_keywords and
                     not any(s_lower.startswith(kw) for kw in code_end_keywords))
                )

                if is_block_end:
                    begin_depth -= 1
                    if self.debug:
                        print(f"[DEBUG] END detected (depth={begin_depth}): {s[:50]}")
                    continue

            # Skip everything inside BEGIN...End blocks
            if begin_depth > 0:
                continue

            # Remove standalone MultiUse lines (outside blocks)
            if s_lower.startswith('multiuse'):
                continue

            # Handle Attribute lines
            if s.startswith('Attribute '):
                if keep_vb_name and 'VB_Name' in line:
                    filtered_lines.append(line)
                continue

            # Keep all other lines (actual code)
            filtered_lines.append(line)

        if self.debug and begin_depth != 0:
            print(f"[DEBUG] Warning: Unbalanced BEGIN/End blocks (final depth={begin_depth})")

        return '\n'.join(filtered_lines)


    def _prompt_overwrite(self, module_name, file_path, comp, edit_mode=False):
        """Compare module content, ignoring ALL Attribute differences for comparison"""
        if self.debug:
            print(f"[DEBUG] Overwrite prompt called, edit_mode={edit_mode}")
        if edit_mode:
            return True  # Always overwrite in edit mode, don't prompt

        file_code = self._read_module_code(file_path)
        visio_code = comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines) if comp.CodeModule.CountOfLines > 0 else ""

        # Normalize both: strip ALL headers for fair comparison
        file_normalized = self._strip_vba_header(file_code, keep_vb_name=False)
        visio_normalized = self._strip_vba_header(visio_code, keep_vb_name=False)

        if file_normalized.strip() == visio_normalized.strip() or self.always_yes:
            return True

        print(f"\n⚠️  Module '{module_name}' differs from Visio. See diff below:")
        for line in unified_diff(
            visio_normalized.splitlines(),
            file_normalized.splitlines(),
            fromfile='Visio',
            tofile='Disk',
            lineterm=''
        ):
            print(line)

        print(f"Overwrite module '{module_name}' in Visio with disk version? (y/N): ", end="")
        ans = input().strip().lower()
        return ans in ("y", "yes")



