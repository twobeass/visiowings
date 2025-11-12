"""VBA Module Import functionality
Document module overwrite logic (force option)
Removes VBA header when importing via force
"""
import win32com.client
from pathlib import Path
import re

class VisioVBAImporter:
    """Importiert VBA-Module in Visio-Dokumente, optional mit force für Document-Module"""
    
    def __init__(self, visio_file_path, force_document=False, debug=False):
        self.visio_file_path = visio_file_path
        self.visio_app = None
        self.doc = None
        self.force_document = force_document
        self.debug = debug

    def connect_to_visio(self):
        """Verbindet sich mit bereits geöffnetem Dokument"""
        try:
            self.visio_app = win32com.client.Dispatch("Visio.Application")
            
            # Find already opened document
            for doc in self.visio_app.Documents:
                if doc.FullName.lower() == str(self.visio_file_path).lower():
                    self.doc = doc
                    if self.debug:
                        print(f"[DEBUG] Verbunden mit Dokument: {doc.Name}")
                    return True
            
            print(f"⚠️  Dokument nicht geöffnet: {self.visio_file_path}")
            print("   Bitte öffne das Dokument in Visio.")
            self.doc = None
            return False
        
        except Exception as e:
            print(f"❌ Fehler beim Verbinden: {e}")
            if self.debug:
                import traceback
                traceback.print_exc()
            self.doc = None
            return False

    def _ensure_connection(self):
        """Stellt sicher, dass die Verbindung zum Dokument noch aktiv ist"""
        try:
            # Test if document is still accessible
            _ = self.doc.Name
            if self.debug:
                print("[DEBUG] Verbindung OK")
            return True
        except Exception:
            if self.debug:
                print("[DEBUG] Verbindung verloren, versuche neu zu verbinden...")
            else:
                print("🔄 Verbindung verloren, versuche neu zu verbinden...")
            return self.connect_to_visio()
    
    def _strip_vba_header(self, code):
        """Entfernt den VBA-Header aus exportierten .cls/.bas-Dateien"""
        lines = code.splitlines()
        code_start = 0
        
        # VBA metadata patterns
        vba_header_pattern = re.compile(r'^(VERSION|Begin|End|Attribute |MultiUse)')
        
        # Find first line of actual code
        for i, line in enumerate(lines):
            stripped = line.strip()
            
            # Skip empty lines and comments at the start
            if not stripped or stripped.startswith("'"):
                continue
            
            # Skip VBA metadata lines
            if vba_header_pattern.match(line):
                continue
            
            # Found actual code
            code_start = i
            break
        
        result = '\n'.join(lines[code_start:])
        
        if self.debug and code_start > 0:
            print(f"[DEBUG] {code_start} Header-Zeilen beim Import entfernt")
        
        return result

    def import_module(self, file_path):
        """Importiert ein einzelnes VBA-Modul, überschreibt Document-Module falls 'force'"""
        
        # Verbindung vor jedem Import prüfen
        if not self._ensure_connection():
            print("⚠️  Keine Verbindung zu Visio - stelle sicher, dass das Dokument geöffnet ist")
            return False
        
        try:
            vb_project = self.doc.VBProject
            file_path = Path(file_path)
            module_name = file_path.stem
            
            if self.debug:
                print(f"[DEBUG] Importiere: {file_path.name}")
            
            # Find existing component with same name
            component = None
            for comp in vb_project.VBComponents:
                if comp.Name == module_name:
                    component = comp
                    break
            
            # Document Module handling (Type 100)
            if component and component.Type == 100:  # vbext_ct_Document
                if self.force_document:
                    code = file_path.read_text(encoding="utf-8")
                    code = self._strip_vba_header(code)
                    
                    cm = component.CodeModule
                    
                    # Delete all existing lines
                    if cm.CountOfLines > 0:
                        cm.DeleteLines(1, cm.CountOfLines)
                    
                    # Add new code
                    if code.strip():  # Only add if there's actual code
                        cm.AddFromString(code)
                    
                    print(f"✓ Code von {file_path.name} (ohne Header) überschrieben (force)")
                    return True
                else:
                    print(f"⚠️  Document-Module '{module_name}' ohne --force übersprungen.")
                    if self.debug:
                        print("[DEBUG] Verwende --force um Document-Module zu überschreiben")
                    return False
            
            # Standard module or class module (Type 1, 2, 3)
            if component:
                if self.debug:
                    print(f"[DEBUG] Entferne existierendes Modul: {module_name}")
                vb_project.VBComponents.Remove(component)
            
            # Import the module
            vb_project.VBComponents.Import(str(file_path))
            print(f"✓ Importiert: {file_path.name}")
            return True
        
        except Exception as e:
            print(f"✗ Fehler beim Importieren von {file_path.name}: {e}")
            if self.debug:
                import traceback
                traceback.print_exc()
            return False
