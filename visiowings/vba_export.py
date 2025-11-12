"""VBA Module Export functionality: improved header stripping and hash-based change detection"""
import win32com.client
import os
from pathlib import Path
import re
import hashlib

class VisioVBAExporter:
    def __init__(self, visio_file_path):
        self.visio_file_path = visio_file_path
        self.visio_app = None
        self.doc = None
    def connect_to_visio(self):
        try:
            self.visio_app = win32com.client.Dispatch("Visio.Application")
            self.doc = self.visio_app.Documents.Open(self.visio_file_path)
            return True
        except Exception as e:
            print(f"❌ Fehler beim Verbinden mit Visio: {e}")
            return False
    def _strip_vba_header_file(self, file_path):
        try:
            text = Path(file_path).read_text(encoding="utf-8")
            header_markers = ['Option Explicit', 'Sub ', 'Function ', 'Property ']
            lines = text.splitlines()
            code_start = 0
            for i, line in enumerate(lines):
                if any(marker in line for marker in header_markers):
                    code_start = i
                    break
            new_text = '\n'.join(lines[code_start:])
            Path(file_path).write_text(new_text, encoding="utf-8")
        except Exception:
            pass
    def _module_content_hash(self, vb_project):
        """Generate content hash of all modules for change detection"""
        try:
            code_parts = []
            for comp in vb_project.VBComponents:
                cm = comp.CodeModule
                # Only hash actual code, not headers
                code = cm.Lines(1, cm.CountOfLines)
                code_parts.append(f"{comp.Name}:{code}")
            return hashlib.md5(''.join(code_parts).encode()).hexdigest()
        except:
            return None
    def export_modules(self, output_dir, last_hash=None):
        if not self.doc:
            print("❌ Kein Dokument geöffnet")
            return []
        try:
            vb_project = self.doc.VBProject
            current_hash = self._module_content_hash(vb_project)
            if last_hash and last_hash == current_hash:
                print("✓ Keine Änderungen erkannt – kein Export notwendig")
                return []
            output_path = Path(output_dir)
            output_path.mkdir(exist_ok=True)
            exported_files = []
            for component in vb_project.VBComponents:
                ext_map = {
                    1: '.bas',
                    2: '.cls',
                    3: '.frm',
                    100: '.cls'
                }
                ext = ext_map.get(component.Type, '.bas')
                file_name = f"{component.Name}{ext}"
                file_path = output_path / file_name
                component.Export(str(file_path))
                # Remove VBA header aggressively
                if component.Type in [1,2,100]:
                    self._strip_vba_header_file(file_path)
                exported_files.append(file_path)
                print(f"✓ Exportiert: {file_name}")
            return exported_files, current_hash
        except Exception as e:
            print(f"❌ Fehler beim Exportieren: {e}")
            print("")
            print("⚠️  Stelle sicher, dass in Visio folgende Einstellung aktiviert ist:")
            print("   Datei → Optionen → Trust Center → Trust Center-Einstellungen")
            print("   → Makroeinstellungen → 'Zugriff auf VBA-Projektobjektmodell vertrauen'")
            return [], None
