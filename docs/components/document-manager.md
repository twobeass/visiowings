# Document Manager Documentation

**File:** `visiowings/document_manager.py`

## Overview

The Document Manager is responsible for discovering and managing multiple Visio documents (drawings, stencils, templates) within a single editing session. It provides a unified interface for accessing VBA projects across all open documents.

## Core Classes

### `VisioDocumentInfo`

Data class representing a single Visio document.

**Attributes:**
```python
class VisioDocumentInfo:
    doc: COM Object          # win32com.client Visio.Document
    name: str                # Display name (e.g., "MyDrawing")
    doc_type: str            # "Drawing" | "Stencil" | "Template"
    folder_name: str         # Folder name for exports (e.g., "mydrawing")
    full_name: str           # Full file path
    has_vba: bool            # True if document has VBA project
```

**Example:**
```python
VisioDocumentInfo(
    doc=<COM Object>,
    name="MyDrawing.vsdm",
    doc_type="Drawing",
    folder_name="mydrawing",
    full_name="C:\\Projects\\MyDrawing.vsdm",
    has_vba=True
)
```

### `VisioDocumentManager`

Manages discovery and access to all relevant Visio documents.

**Constructor:**
```python
VisioDocumentManager(visio_file_path: str, debug: bool = False)
```

**Key Methods:**
- `connect_to_visio()` - Establish COM connection
- `get_all_documents_with_vba()` - List all documents with VBA
- `get_main_document()` - Get primary document
- `print_summary()` - Display discovered documents

## Document Discovery

### Algorithm

1. **Connect to Visio Application**
   ```python
   self.visio_app = win32com.client.GetActiveObject("Visio.Application")
   ```

2. **Find Main Document**
   - Iterate through `visio_app.Documents`
   - Match by `FullName` or `Name` against `visio_file_path`

3. **Discover Referenced Documents**
   - Check main document's stencils
   - Check templates
   - Identify related drawings

4. **Filter VBA Projects**
   - Test each document: `doc.HasVBProject`
   - Skip documents without VBA

### Document Type Detection

```python
def _get_document_type(doc) -> str:
    try:
        if doc.Type == 1:  # visDrawing
            return "Drawing"
        elif doc.Type == 2:  # visStencil
            return "Stencil"
        elif doc.Type == 3:  # visTemplate
            return "Template"
    except:
        pass
    return "Unknown"
```

### Folder Name Generation

```python
def _generate_folder_name(doc_name: str) -> str:
    # Remove extension
    base_name = Path(doc_name).stem
    
    # Sanitize for file system
    folder_name = base_name.lower()
    folder_name = folder_name.replace(' ', '_')
    folder_name = ''.join(c for c in folder_name if c.isalnum() or c == '_')
    
    return folder_name
```

**Examples:**
- `"My Drawing.vsdm"` → `"my_drawing"`
- `"Company-Stencil.vssx"` → `"companystencil"`
- `"Template 2.0.vstx"` → `"template_20"`

## Multi-Document Scenarios

### Scenario 1: Drawing with Custom Stencil

**Open Documents:**
- `Project.vsdm` (main drawing)
- `CustomShapes.vssm` (stencil with VBA)

**Export Structure:**
```
vba_modules/
├── project/
│   ├── Module1.bas
│   └── ThisDocument.cls
└── customshapes/
    ├── ShapeHelper.bas
    └── ThisDocument.cls
```

### Scenario 2: Template Development

**Open Documents:**
- `MyTemplate.vstx` (template)
- `TemplateStencil.vssx` (stencil)

**Export Structure:**
```
vba_modules/
├── mytemplate/
│   └── TemplateInit.bas
└── templatestencil/
    └── StencilCode.bas
```

## VBA Project Access

### Checking VBA Presence

```python
def has_vba_project(doc) -> bool:
    try:
        vb_project = doc.VBProject
        if vb_project.VBComponents.Count > 0:
            return True
    except:
        pass
    return False
```

### Accessing Components

```python
for doc_info in manager.get_all_documents_with_vba():
    vb_project = doc_info.doc.VBProject
    
    for component in vb_project.VBComponents:
        print(f"{doc_info.name}: {component.Name}")
```

## Summary Output

### Single Document
```
📝 Document: MyDrawing.vsdm
   Type: Drawing
   VBA: Yes
```

### Multiple Documents
```
📝 Main Document: MyDrawing.vsdm
   Type: Drawing
   VBA: Yes

📄 Additional Documents:
   • CustomStencil.vssm (Stencil, VBA: Yes)
   • SharedTemplate.vstx (Template, VBA: No - skipped)

ℹ️  Total documents with VBA: 2
```

## Error Handling

### Connection Failures

```python
def connect_to_visio(self) -> bool:
    try:
        self.visio_app = win32com.client.GetActiveObject("Visio.Application")
        # Find main document...
        return True
    except:
        print("❌ Visio not running or document not open")
        return False
```

### Document Access Errors

```python
try:
    vb_project = doc.VBProject
except:
    # Missing VBA access permission
    print("⚠️  Enable 'Trust access to VBA project object model'")
```

## Performance Considerations

### Lazy Discovery

Documents are discovered once during initialization:
```python
def __init__(self, visio_file_path, debug=False):
    self._documents_cache = None  # Populated on first access
```

### COM Object Reuse

Document COM objects are reused throughout the session:
```python
# Store references, don't recreate
self.all_docs = self._discover_documents()
```

## Thread Safety

Document Manager is **not thread-safe**. Create separate instances for different threads:

```python
# Main thread
exporter_manager = VisioDocumentManager(file_path)

# Watcher thread
importer_manager = VisioDocumentManager(file_path)
```

## Debugging

With `debug=True`:
```
[DEBUG] Connecting to Visio...
[DEBUG] Found main document: MyDrawing.vsdm
[DEBUG] Scanning for additional documents...
[DEBUG] Found stencil: CustomStencil.vssm
[DEBUG] Document map created: ['mydrawing', 'customstencil']
```

## Best Practices

### Always Check Connection
```python
manager = VisioDocumentManager(file_path)
if not manager.connect_to_visio():
    return  # Handle error
```

### Cache Document List
```python
# Good - cache results
docs_with_vba = manager.get_all_documents_with_vba()
for doc_info in docs_with_vba:
    # process...

# Avoid - repeated discovery
for doc_info in manager.get_all_documents_with_vba():  # Repeated calls
    # ...
```

### Handle Missing Documents Gracefully
```python
docs = manager.get_all_documents_with_vba()
if not docs:
    print("⚠️  No documents with VBA found")
    return
```
