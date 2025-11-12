"""VBA Module Export functionality: strips VBA header when exporting Document/Class Modules"""
import win32com.client
import os
from pathlib import Path
import re

class VisioVBAExporter:
    """Exportiert VBA-Module aus Visio-Dokumenten"""
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
        """Strip VBA header from file exported by Visio/Excel for .cls/.bas/.frm files"""
        try:
            text = Path(file_path).read_text(encoding="utf-8")
            lines = text.splitlines()
            code_start = 0
            header_pattern = re.compile(r'^(VERSION|Begin|End|Attribute )')
            for i, line in enumerate(lines):
                if not header_pattern.match(line):
                    code_start = i
                    break
            new_text = '\n'.join(lines[code_start:])
            Path(file_path).write_text(new_text, encoding="utf-8")
        except Exception:
            pass
    def export_modules(self, output_dir):
        if not self.doc:
            print("❌ Kein Dokument geöffnet")
            return []
        try:
            vb_project = self.doc.VBProject
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
                # Exportiere das Modul
                component.Export(str(file_path))
                # Entferne Header bei Class und Document
                if component.Type in [2, 100]:
                    self._strip_vba_header_file(file_path)
                exported_files.append(file_path)
                print(f"✓ Exportiert: {file_name}")
            return exported_files
        except Exception as e:
            print(f"❌ Fehler beim Exportieren: {e}")
            print("")
            print("⚠️  Stelle sicher, dass in Visio folgende Einstellung aktiviert ist:")
            print("   Datei → Optionen → Trust Center → Trust Center-Einstellungen")
            print("   → Makroeinstellungen → 'Zugriff auf VBA-Projektobjektmodell vertrauen'")
            return []
