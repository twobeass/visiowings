# Multi-Document Support for visiowings

## 🆕 Feature Branch: `feature/multi-document-support`

### 🎯 Ziel

Unterstützung für Visio-Zeichnungen (`.vsdx`) mit Schablonen (`.vssm`), die VBA-Code enthalten.

## Problem & Lösung

### Problem

Bisher:
- visiowings funktioniert nur mit `.vsdm` Dateien (Makro-aktivierte Zeichnungen)
- VBA-Code muss im Hauptdokument sein
- **Neues Szenario**: `.vsdx` Zeichnung + `.vssm` Schablone mit VBA-Code

### Lösung: Auto-Detection

1. **Automatische Erkennung** aller geöffneten Dokumente mit VBA-Code
2. **Separate Unterordner** pro Dokument:
   ```
   project/
   ├── drawing.vsdx (geöffnet)
   ├── mystencil.vssm (geöffnet)
   └── vba_modules/
       ├── drawing/          # Code aus drawing.vsdx
       │   └── Module1.bas
       └── mystencil/        # Code aus mystencil.vssm
           ├── ThisDocument.cls
           ├── StencilModule1.bas
           └── StencilClass1.cls
   ```

3. **Automatische Zuordnung** beim Import: Ordnername → Dokument

## Neue Komponenten

### 1. `document_manager.py`

**VisioDocumentManager** - Verwaltet mehrere Visio-Dokumente:

```python
from visiowings.document_manager import VisioDocumentManager

manager = VisioDocumentManager("drawing.vsdx", debug=True)
if manager.connect_to_visio():
    # Alle Dokumente mit VBA-Code
    for doc_info in manager.get_all_documents_with_vba():
        print(f"{doc_info.name} ({doc_info.get_type_name()})")
    
    # Nur Schablonen
    stencils = manager.get_stencils()
```

**VisioDocumentInfo** - Informationen über ein Dokument:
- `name`: Dokumentname
- `type`: Dokumenttyp (Drawing=1, Stencil=2, Template=3)
- `has_vba`: Hat VBA-Code?
- `folder_name`: Bereinigter Name für Ordner

### 2. Erweiterte `vba_export.py`

**Neue Rückgabewerte**:
```python
# Alt (Single Document):
exported_files, hash = exporter.export_modules(output_dir)

# Neu (Multi Document):
all_exported, all_hashes = exporter.export_modules(output_dir)
# all_exported = {"drawing": [file1, file2], "mystencil": [file3]}
# all_hashes = {"drawing": "abc123...", "mystencil": "def456..."}
```

**Hash-Tracking pro Dokument**:
- Jedes Dokument hat eigenen Hash
- Export nur wenn Dokument sich geändert hat

### 3. Erweiterte `vba_import.py`

**Automatische Dokument-Zuordnung**:
```python
# Datei: vba_modules/mystencil/Module1.bas
# → Wird automatisch in "mystencil.vssm" importiert

importer.import_module(Path("vba_modules/mystencil/Module1.bas"))
```

**Backward Compatibility**:
- Dateien im Root-Verzeichnis → Hauptdokument
- Dateien in Unterordnern → Entsprechendes Dokument

### 4. Erweiterte `file_watcher.py`

**Rekursive Überwachung**:
```python
observer.schedule(
    event_handler,
    watch_directory,
    recursive=True  # Jetzt aktiviert!
)
```

**Multi-Document Hash-Tracking**:
```python
self.last_export_hashes = {
    "drawing": "abc123...",
    "mystencil": "def456..."
}
```

## Verwendung

### Voraussetzungen

1. **Öffne alle Dokumente in Visio**:
   - Hauptzeichnung: `drawing.vsdx`
   - Schablone(n): `mystencil.vssm`

2. **Stelle sicher, dass VBA-Code vorhanden ist**:
   - In Visio: Alt+F11 → VBA-Editor
   - Schablone muss VBA-Module enthalten

### Beispiel-Workflow

```bash
# 1. Öffne drawing.vsdx in Visio
#    Dies lädt auch die referenzierte mystencil.vssm

# 2. Starte visiowings
cd C:/Projects/MyVisioProject
visiowings edit --file "drawing.vsdx" --force --bidirectional

# Output:
# 📂 Visio-Datei: C:\Projects\MyVisioProject\drawing.vsdx
# 📁 Export-Verzeichnis: C:\Projects\MyVisioProject
#
# === Exportiere VBA-Module ===
#
# 📚 Gefundene Dokumente mit VBA: 2
#    📄 Hauptdokument: drawing.vsdx (Drawing)
#    📋 Schablonen: 1
#       - mystencil.vssm
#
# ✓ Exportiert: drawing/Module1.bas
# ✓ Exportiert: mystencil/ThisDocument.cls
# ✓ Exportiert: mystencil/StencilModule1.bas
#
# ✓ 3 Module aus 2 Dokumenten exportiert
#
# === Starte Live-Synchronisation ===
# 👁️  Überwache Verzeichnis: C:\Projects\MyVisioProject
# 💾 Speichere Dateien in VS Code (Ctrl+S) um sie nach Visio zu synchronisieren
# 🔄 Bidirektionaler Sync: Änderungen in Visio werden automatisch nach VSCode exportiert.
# ⏸️  Drücke Ctrl+C zum Beenden...

# 3. Bearbeite in VS Code
code .
# Ändere: vba_modules/mystencil/StencilModule1.bas
# Speichern (Ctrl+S)

# Output:
# 📝 Änderung erkannt: mystencil/StencilModule1.bas
# ✓ Importiert: mystencil/StencilModule1.bas

# 4. Bearbeite in Visio VBA-Editor (Alt+F11)
# Ändere Code in mystencil.vssm
# Nach ~4 Sekunden:

# Output:
# 🔄 Visio-Dokument(e) wurden synchronisiert → VSCode.
```

## Ordnerstruktur

### Beispiel 1: Zeichnung + 1 Schablone

```
project/
├── drawing.vsdx
├── mystencil.vssm
└── vba_modules/
    ├── drawing/
    │   ├── Module1.bas
    │   └── ClassModule1.cls
    └── mystencil/
        ├── ThisDocument.cls
        ├── StencilModule1.bas
        └── StencilClass1.cls
```

### Beispiel 2: Zeichnung + Mehrere Schablonen

```
project/
├── drawing.vsdx
├── shapes.vssm
├── utilities.vssm
└── vba_modules/
    ├── drawing/
    │   └── Module1.bas
    ├── shapes/
    │   ├── ThisDocument.cls
    │   └── ShapeHelpers.bas
    └── utilities/
        ├── ThisDocument.cls
        └── UtilityFunctions.bas
```

### Beispiel 3: Nur Makro-aktivierte Zeichnung (Backward Compatible)

```
project/
├── document.vsdm
└── vba_modules/
    └── document/
        ├── ThisDocument.cls
        ├── Module1.bas
        └── ClassModule1.cls
```

## Debug-Modus

Für detaillierte Informationen:

```bash
visiowings edit --file "drawing.vsdx" --force --bidirectional --debug
```

**Zusätzliche Ausgaben**:
```
[DEBUG] Gefundene Dokumente: 2
[DEBUG]   - VisioDocumentInfo(name='drawing.vsdx', type=Drawing, has_vba=True)
[DEBUG]   - VisioDocumentInfo(name='mystencil.vssm', type=Stencil, has_vba=True)
[DEBUG] VBA gefunden in: drawing.vsdx (Drawing)
[DEBUG] VBA gefunden in: mystencil.vssm (Stencil)
[DEBUG] Dokument-Map erstellt: ['drawing', 'mystencil']
[DEBUG] Exportiere drawing.vsdx...
[DEBUG] Hash berechnet: abc123... (2 Module)
[DEBUG] Exportiere mystencil.vssm...
[DEBUG] Hash berechnet: def456... (3 Module)
[DEBUG] drawing: Hash abc123...
[DEBUG] mystencil: Hash def456...
```

## Technische Details

### Dokument-Typen

```python
class VisioDocumentType:
    DRAWING = 1    # visTypeDrawing - .vsdx, .vsdm
    STENCIL = 2    # visTypeStencil - .vssx, .vssm
    TEMPLATE = 3   # visTypeTemplate - .vstx, .vstm
```

### Ordnernamen-Bereinigung

```python
# "My Stencil (2024).vssm" → "my_stencil_2024"
# "Shapes & Utilities.vssm" → "shapes_utilities"
# "Tool-Box.vssm" → "tool_box"
```

Regeln:
- Dateiendung entfernen
- Leerzeichen → Unterstrich
- Sonderzeichen → Unterstrich
- Kleinbuchstaben
- Keine mehrfachen/führenden/abschließenden Unterstriche

### Hash-Berechnung

Pro Dokument:
```python
hash_input = f"{module1_name}:{module1_code}{module2_name}:{module2_code}..."
content_hash = md5(hash_input).hexdigest()
```

## Troubleshooting

### "Keine Dokumente mit VBA-Code gefunden"

**Problem**: Schablone enthält keinen VBA-Code

**Lösung**:
1. Öffne Schablone in Visio
2. Alt+F11 → VBA-Editor
3. Füge mindestens ein Modul hinzu
4. Speichere Schablone als `.vssm`

### "Datei wird falschem Dokument zugeordnet"

**Problem**: Import findet Dokument nicht

**Lösung** (Debug):
```bash
visiowings edit --file "drawing.vsdx" --debug

# Prüfe Output:
# [DEBUG] Dokument-Map erstellt: ['drawing', 'mystencil']
# [DEBUG] Datei Module1.bas gehört zu Dokument: mystencil
```

**Manueller Fix**:
- Verschiebe Datei in korrekten Unterordner
- Ordnername muss mit sanitized document name übereinstimmen

### "Schablone nicht geöffnet"

**Problem**: visiowings findet Schablone nicht

**Lösung**:
1. **Vor** visiowings-Start:
   - Öffne Hauptzeichnung in Visio
   - Öffne Schablone explizit (Datei → Formen → Eigene Formen)
2. **Oder**: Zeichnung referenziert Schablone automatisch

### "Hash-Werte stimmen nicht"

**Problem**: Export wird trotz identischem Code getriggert

**Debug**:
```bash
visiowings edit --file "drawing.vsdx" --bidirectional --debug

# Prüfe Output:
# [DEBUG] mystencil: Last hash: abc123...
# [DEBUG] mystencil: Current hash: abc123...
# [DEBUG] mystencil: Hashes identisch - kein Export
```

Wenn Hashes unterschiedlich obwohl Code gleich:
- Möglicherweise Whitespace-Änderungen
- Visio fügt Kommentare/Metadaten hinzu

## Backward Compatibility

### Single-Document Modus

Falls nur **ein** Dokument VBA-Code hat:
```
project/
├── document.vsdm
└── vba_modules/
    └── document/        # Unterordner wird trotzdem erstellt
        ├── Module1.bas
        └── ClassModule1.cls
```

### Legacy-Struktur (ohne Unterordner)

Falls Dateien direkt in `vba_modules/` liegen:
```
project/
├── document.vsdm
└── vba_modules/
    ├── Module1.bas      # Wird Hauptdokument zugeordnet
    └── ClassModule1.cls  # Wird Hauptdokument zugeordnet
```

→ Import funktioniert, wird automatisch Hauptdokument zugeordnet

## Nächste Schritte

### Testing

1. **Teste mit `.vsdx` + `.vssm`**:
   ```bash
   # Erstelle Test-Setup
   # - drawing.vsdx (mit oder ohne VBA)
   # - mystencil.vssm (mit VBA)
   
   visiowings edit --file "drawing.vsdx" --force --bidirectional --debug
   ```

2. **Teste Backward Compatibility**:
   ```bash
   # Teste mit bestehendem .vsdm Projekt
   visiowings edit --file "old_document.vsdm" --force --bidirectional
   ```

3. **Teste Multi-Stencil**:
   ```bash
   # Öffne mehrere Schablonen in Visio
   visiowings edit --file "drawing.vsdx" --force --bidirectional --debug
   ```

### Merge nach Main

Nach erfolgreichem Testing:
```bash
git checkout main
git merge feature/multi-document-support
git push origin main
```

## Weitere Features (Optional)

### Geplante Erweiterungen

- [ ] `.visiowingsignore` für Dokument-Filter
- [ ] `--document` Flag für explizite Auswahl
- [ ] Konfigurierbares Polling-Intervall pro Dokument
- [ ] Dokumenten-Status in CLI anzeigen
- [ ] Warnung wenn Schablone geändert aber nicht gespeichert

---

## Zusammenfassung

✅ **Automatische Erkennung** aller Dokumente mit VBA
✅ **Separate Ordner** pro Dokument
✅ **Automatische Zuordnung** beim Import
✅ **Hash-Tracking** pro Dokument
✅ **Rekursive Überwachung** aller Unterordner
✅ **Backward Compatible** mit Single-Document
✅ **Debug-Modus** für Troubleshooting

**Use Case erfüllt**: `.vsdx` mit `.vssm` Schablonen! 🎉
