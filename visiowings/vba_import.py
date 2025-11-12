"""VBA Module Import functionality
Document module overwrite logic (force option)"""
import win32com.client
from pathlib import Path

class VisioVBAImporter:
    """Importiert VBA-Module in Visio-Dokumente, optional mit force für Document-Module"""
    def __init__(self, visio_file_path, force_document=False):
        self.visio_file_path = visio_file_path
        self.visio_app = None
        self.doc = None
        self.force_document = force_document

    def connect_to_visio(self):
        """Verbindet sich mit bereits geöffnetem Dokument"""
        try:
            self.visio_app = win32com.client.Dispatch("Visio.Application")
            for doc in self.visio_app.Documents:
                if doc.FullName.lower() == str(self.visio_file_path).lower():
                    self.doc = doc
                    return True
            print(f"⚠️  Dokument nicht geöffnet: {self.visio_file_path}")
            print("   Bitte öffne das Dokument in Visio.")
            return False
        except Exception as e:
            print(f"❌ Fehler beim Verbinden: {e}")
            return False

    def import_module(self, file_path):
        """Importiert ein einzelnes VBA-Modul, überschreibt Document-Module falls 'force'"""
        if not self.doc:
            return False
        try:
            vb_project = self.doc.VBProject
            file_path = Path(file_path)
            module_name = file_path.stem
            # Suche Component
            component = None
            for comp in vb_project.VBComponents:
                if comp.Name == module_name:
                    component = comp
                    break
            # Document Module handling
            if component and component.Type == 100:  # vbext_ct_Document
                if self.force_document:
                    code = file_path.read_text(encoding="utf-8")
                    cm = component.CodeModule
                    cm.DeleteLines(1, cm.CountOfLines)
                    cm.AddFromString(code)
                    print(f"✓ Code von {file_path.name} überschrieben (force)")
                    return True
                else:
                    print(f"⚠️  Document-Module '{module_name}' ohne --force übersprungen.")
                    return False
            # Standardmodul oder Klassenmodul
            if component:
                vb_project.VBComponents.Remove(component)
            vb_project.VBComponents.Import(str(file_path))
            print(f"✓ Importiert: {file_path.name}")
            return True
        except Exception as e:
            print(f"✗ Fehler beim Importieren von {file_path.name}: {e}")
            return False
