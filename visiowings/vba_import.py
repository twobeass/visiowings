import hashlib
import os
import re
from difflib import unified_diff
from pathlib import Path

import pythoncom

from .document_manager import VisioDocumentManager, sanitize_document_name
from .encoding import DEFAULT_CODEPAGE, resolve_encoding


class VisioVBAImporter:
    def __init__(self, visio_file_path, force_document=False, debug=False, silent_reconnect=False, always_yes=False, user_codepage=None, use_rubberduck=False):
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
        self.use_rubberduck = use_rubberduck

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
        Returns the Visio document object associated with this file, based on folder structure.
        Searches upwards to find a directory matching a known document.
        """
        current_path = file_path.parent
        # Traverse up to find a matching document root folder
        # Limit traversal to avoid going too far up (e.g. to root /)
        # We assume the input_dir is the root of the operation.
        # But we don't have input_dir easily here. We can assume we stop when we hit a known doc folder.

        # Try direct parent first
        # Try direct parent first
        sanitized_parent = sanitize_document_name(current_path.name)
        if sanitized_parent in self.document_map:
            if self.debug:
                print(f"[DEBUG] File {file_path.name} belongs to document: {sanitized_parent}")
            return self.document_map[sanitized_parent]

        # In rubberduck mode, we might be deep in subfolders
        if self.use_rubberduck:
            # Try walking up
            for _ in range(10): # Max depth safety
                sanitized_current = sanitize_document_name(current_path.name)
                if sanitized_current in self.document_map:
                    if self.debug:
                        print(f"[DEBUG] File {file_path.name} (nested) belongs to document: {sanitized_current}")
                    return self.document_map[sanitized_current]
                if current_path.parent == current_path: # Root
                    break
                current_path = current_path.parent
            
            # If we are here in RD mode, we found NO match.
            # Fallback to main document is DANGEROUS in multi-file projects.
            if self.debug:
                print(f"[DEBUG] No document match found for {file_path} in RD mode. Skipping main doc fallback.")
            return None

        # Fallback for root files (legacy single-document support)
        if self.debug:
            print(f"[DEBUG] No folder match for {file_path.name}, attempting main doc fallback")
        return self.doc_manager.get_main_document()

    def _ensure_folder_annotation(self, content, file_path, doc_info):
        """Inject or update Rubberduck @Folder annotation based on file path"""
        if not self.use_rubberduck:
            return content

        # Calculate relative path from document root
        try:
            # We need to find the document root directory in the path
            # This is tricky because file_path is absolute.
            # We can use the doc_info.folder_name to find the split point.
            parts = file_path.parts
            if doc_info.folder_name in parts:
                idx = parts.index(doc_info.folder_name)
                # Subfolder parts come after the document folder
                sub_parts = parts[idx+1:-1] # -1 to exclude filename
                if not sub_parts:
                     return content # Root of document, no folder annotation needed

                # Inject as comment: '@Folder("Path")
                folder_annotation = f"'@Folder(\"{ '.'.join(sub_parts) }\")"

                # Check if annotation exists (any variant)
                if "@Folder" in content:
                    # Update existing (regex replace), handling optional comment prefix in existing file
                    # We standardize it to have the comment prefix
                    content = re.sub(r"(')?\s*@Folder\s*\(\s*\"[^\"]+\"\s*\)", folder_annotation, content, count=1)
                else:
                    # Inject
                    # Preferred location: Top of file, but after VB_Name if present.
                    # Or before Option Explicit.
                    lines = content.splitlines()
                    insert_idx = 0

                    # Skip Attribute lines at top
                    while insert_idx < len(lines) and lines[insert_idx].strip().startswith("Attribute "):
                         insert_idx += 1

                    # Check for Option Explicit
                    option_explicit_idx = -1
                    for i, line in enumerate(lines):
                        if line.strip().lower() == "option explicit":
                            option_explicit_idx = i
                            break

                    if option_explicit_idx != -1:
                        # Insert before Option Explicit
                        lines.insert(option_explicit_idx, folder_annotation)
                    else:
                        # Insert after attributes
                        lines.insert(insert_idx, folder_annotation)

                    content = "\n".join(lines)

        except Exception as e:
            if self.debug:
                print(f"[DEBUG] Error ensuring folder annotation: {e}")

        return content

    def _create_temp_codepage_file(self, file_path, codepage, doc_info=None):
        """Create a temporary file with configured encoding for VBA import."""
        import tempfile
        try:
            text = file_path.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            text = file_path.read_text(encoding=codepage)

        # Handle Rubberduck annotations
        if self.use_rubberduck and doc_info:
            text = self._ensure_folder_annotation(text, file_path, doc_info)

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
        """Import a single module. Used by file watcher."""
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

            # Special handling for Document modules
            if component and component.Type == 100:
                if self.force_document:
                    self._import_document_module_content(component, file_path)
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

            temp_file = self._create_temp_codepage_file(file_path, self.codepage, doc_info=target_doc_info)
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

    def _compare_module_content(self, file_path, component):
        """Compare local file with Visio module content using normalization.
        Returns: (are_different, local_hash, visio_hash)
        """
        try:
            file_code = self._read_module_code(file_path)
            visio_code = component.CodeModule.Lines(1, component.CodeModule.CountOfLines) if component.CodeModule.CountOfLines > 0 else ""

            # Normalize both: strip ALL headers for fair comparison
            file_normalized = self._strip_vba_header(file_code, keep_vb_name=False)
            visio_normalized = self._strip_vba_header(visio_code, keep_vb_name=False)

            # Further normalize whitespace
            file_final = self._normalize_content(file_normalized)
            visio_final = self._normalize_content(visio_normalized)

            # Calculate hashes
            local_hash = hashlib.md5(file_final.encode()).hexdigest()[:8]
            visio_hash = hashlib.md5(visio_final.encode()).hexdigest()[:8]

            are_different = file_final != visio_final

            return are_different, local_hash, visio_hash
        except Exception as e:
            if self.debug:
                print(f"[DEBUG] Error comparing {file_path.name}: {type(e).__name__}: {e}")
            return True, None, None

    def _prompt_overwrite(self, module_name, file_path, comp, edit_mode=False):
        """Compare module content, ignoring ALL Attribute differences for comparison"""
        if self.debug:
            print(f"[DEBUG] Overwrite prompt called, edit_mode={edit_mode}")
        if edit_mode:
            return True  # Always overwrite in edit mode, don't prompt

        are_different, _, _ = self._compare_module_content(file_path, comp)

        if not are_different or self.always_yes:
            return True

        print(f"\n⚠️  Module '{module_name}' differs from Visio. See diff below:")

        # Show nice diff
        file_code = self._read_module_code(file_path)
        visio_code = comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines) if comp.CodeModule.CountOfLines > 0 else ""
        file_normalized = self._strip_vba_header(file_code, keep_vb_name=False)
        visio_normalized = self._strip_vba_header(visio_code, keep_vb_name=False)

        for line in unified_diff(
            visio_normalized.splitlines(),
            file_normalized.splitlines(),
            fromfile='Visio',
            tofile='Disk',
            lineterm=''
        ):
            print(line)

        print(f"Overwrite module '{module_name}' in Visio with disk version? (y/N/a): ", end="")
        ans = input().strip().lower()

        if ans in ("a", "all"):
            self.always_yes = True
            return True

        return ans in ("y", "yes")

    def _import_document_module_content(self, component, file_path):
        """Helper to overwrite document module content"""
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

    def import_modules_from_dir(self, input_dir):
        """Batch import all modules from directory with conflict checking"""
        if not self.connect_to_visio():
            return 0

        input_dir = Path(input_dir)
        documents_to_process = {}

        # 1. Discovery Phase
        # Find all candidate files
        candidate_files = []
        # In RD mode, we need recursive search
        if self.use_rubberduck:
            for ext in ['*.bas', '*.cls', '*.frm']:
                candidate_files.extend(input_dir.rglob(ext))
        else:
            for ext in ['*.bas', '*.cls', '*.frm']:
                candidate_files.extend(input_dir.glob(ext)) # Root files
                for doc_folder in self.get_document_folders():
                    candidate_files.extend((input_dir / doc_folder).glob(ext)) # Subdir files

        # Map files to documents
        for file_path in candidate_files:
            doc_info = self._find_document_for_file(file_path)
            if doc_info:
                if doc_info.folder_name not in documents_to_process:
                    documents_to_process[doc_info.folder_name] = {
                        'doc_info': doc_info,
                        'files': []
                    }
                documents_to_process[doc_info.folder_name]['files'].append(file_path)

        total_imported = 0

        # 2. Process each document
        for doc_folder, data in documents_to_process.items():
            doc_info = data['doc_info']
            files = data['files']
            vb_project = doc_info.doc.VBProject

            files_with_changes = {}
            files_to_import = []
            files_identical_count = 0

            # Check for conflicts
            for file_path in files:
                module_name = file_path.stem
                component = None
                for comp in vb_project.VBComponents:
                    if comp.Name == module_name:
                        component = comp
                        break

                if component:
                    if component.Type == 100: # Document module
                        if self.force_document:
                            files_to_import.append((file_path, component, True)) # True = is_doc_mod
                        else:
                            print(f"⚠️  Document module '{module_name}' skipped without --force.")
                    else:
                        are_different, _, _ = self._compare_module_content(file_path, component)
                        if are_different:
                             files_with_changes[module_name] = {
                                'path': file_path,
                                'component': component
                             }
                        else:
                             # No changes, but we might want to "refresh" it?
                             # Usually if identical, we skip to save time/risk, unless specifically requested?
                             # Export skips identical. Import should likely skip identical too unless we are strictly overwriting.
                             # But let's assume if it's identical we skip it for safety/speed.
                             files_identical_count += 1
                else:
                    # New module
                    files_to_import.append((file_path, None, False))

            # Handle conflicts
            if files_with_changes:
                print(f"\n⚠️  Local changes detected in {doc_info.name} (Importing to Visio):")
                for fname in files_with_changes.keys():
                    print(f"   - {doc_info.folder_name}/{fname}")

                print("\nOptions:")
                print("  o - Overwrite all in Visio with local content")
                print("  s - Skip changed files (keep Visio content)")
                print("  i - Interactive (choose per file)")
                print("  c - Cancel import for this document")

                response = input("\nChoose action (o/s/i/C): ").strip().lower()

                if response == 'o':
                    # Overwrite all
                    print(f"✓ Will overwrite {len(files_with_changes)} file(s)")
                    for fname, info in files_with_changes.items():
                         files_to_import.append((info['path'], info['component'], False))

                elif response == 's':
                    # Skip all
                    print(f"✓ Will skip {len(files_with_changes)} changed file(s)")

                elif response == 'i':
                    # Interactive
                     for fname, info in files_with_changes.items():
                        print(f"\n{doc_info.folder_name}/{fname}")

                        # Show diff
                        file_code = self._read_module_code(info['path'])
                        visio_code = info['component'].CodeModule.Lines(1, info['component'].CodeModule.CountOfLines)

                        file_normalized = self._strip_vba_header(file_code, keep_vb_name=False)
                        visio_normalized = self._strip_vba_header(visio_code, keep_vb_name=False)

                        for line in unified_diff(
                            visio_normalized.splitlines(),
                            file_normalized.splitlines(),
                            fromfile='Visio',
                            tofile='Disk',
                            lineterm=''
                        ):
                            print(line)

                        choice = input("  Overwrite Visio module? (y/N): ").strip().lower()
                        if choice in ('y', 'yes'):
                             files_to_import.append((info['path'], info['component'], False))
                else:
                    print(f"❌ Import cancelled for {doc_info.name}")
                    continue

            # Execute Imports
            for file_path, component, is_doc_mod in files_to_import:
                try:
                    if is_doc_mod:
                        self._import_document_module_content(component, file_path)
                        print(f"✓ Imported: {doc_info.folder_name}/{file_path.name} (force)")
                    else:
                        if component:
                            vb_project.VBComponents.Remove(component)

                        temp_file = self._create_temp_codepage_file(file_path, self.codepage, doc_info=doc_info)
                        vb_project.VBComponents.Import(str(temp_file))

                        # Clean up
                        if temp_file and temp_file != str(file_path):
                            try:
                                os.unlink(temp_file)
                            except Exception:
                                pass

                        print(f"✓ Imported: {doc_info.folder_name}/{file_path.name}")
                    total_imported += 1
                except Exception as e:
                    print(f"✗ Error importing {file_path.name}: {type(e).__name__}: {e}")
            
            if files_identical_count > 0:
                print(f"✓ {files_identical_count} modules up-to-date (skipped)")

        return total_imported
