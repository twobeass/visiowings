# Architecture Overview

## System Design

Visiowings is a Python-based bridge between Microsoft Visio's VBA environment and modern code editors (primarily VS Code). It enables bidirectional synchronization of VBA modules using COM automation and file system watchers.

## Core Components

```
┌─────────────────────────────────────────────────────────────┐
│                          CLI Layer                          │
│                   (cli.py - argparse)                       │
└──────────────┬────────────────────────┬─────────────────────┘
               │                        │
               ▼                        ▼
    ┌──────────────────┐      ┌──────────────────┐
    │  VBA Exporter    │      │  VBA Importer    │
    │  (vba_export.py) │      │  (vba_import.py) │
    └────────┬─────────┘      └─────────┬────────┘
             │                          │
             └────────┬─────────────────┘
                      │
                      ▼
            ┌──────────────────┐
            │ Document Manager │
            │(document_manager)│
            └────────┬─────────┘
                     │
                     ▼
            ┌──────────────────┐
            │ Visio Connection │
            │  (via pywin32)   │
            └──────────────────┘
                     │
                     ▼
            ┌──────────────────┐
            │  Microsoft Visio │
            │   (COM Object)   │
            └──────────────────┘
```

### 1. CLI Layer (`cli.py`)

**Purpose:** Entry point for all user interactions

**Commands:**
- `edit` - Full bidirectional sync with live file watching
- `export` - One-time export of VBA modules from Visio
- `import` - One-time import of VBA modules into Visio

**Responsibilities:**
- Parse command-line arguments
- Orchestrate component initialization
- Coordinate export → watch → import workflow

### 2. Document Manager (`document_manager.py`)

**Purpose:** Centralized management of multiple Visio documents

**Key Features:**
- Discovers all open Visio documents (drawings, stencils, templates)
- Maps documents to folder names for multi-document projects
- Provides VBA project access for each document
- Handles document-specific metadata (type, path, VBA presence)

**Data Structures:**
```python
class VisioDocumentInfo:
    doc: COM Object          # Visio Document
    name: str                # Display name
    doc_type: str            # Drawing/Stencil/Template
    folder_name: str         # Export folder name
    full_name: str           # Full path
    has_vba: bool            # VBA project present
```

### 3. VBA Exporter (`vba_export.py`)

**Purpose:** Extract VBA modules from Visio to file system

**Key Responsibilities:**
- Export VBA components from Visio VBProject
- **Rubberduck Integration:** Parse `@Folder` annotations and map to directory structure
- Strip VBA headers while preserving `Attribute VB_Name`
- Convert encoding from cp1252 (Visio export) to UTF-8 (VS Code)
- Normalize content for comparison (whitespace, line endings)
- Detect local changes and prompt user for conflict resolution
- Calculate MD5 hashes for change detection

**File Processing Pipeline:**
```
Visio VBComponent
       ↓
  Export to temp file (cp1252)
       ↓
  Strip VBA headers
       ↓
  Convert to UTF-8
       ↓
  Compare with local file (if exists)
       ↓
  User decision (overwrite/skip/interactive)
       ↓
  Write to disk
```

### 4. VBA Importer (`vba_import.py`)

**Purpose:** Inject VBA modules from file system into Visio

**Key Responsibilities:**
- **Rubberduck Integration:** Recursively discover files and inject `@Folder` annotations based on path
- Repair and normalize VBA file headers
- Handle encoding conversions (UTF-8 → cp1252 → UTF-8 roundtrip)
- Preserve comments and Option Explicit statements
- Safely overwrite document modules (with `--force` flag)
- Auto-detect target document from folder structure

**Import Pipeline:**
```
Local .bas/.cls/.frm file
       ↓
  Read as UTF-8
       ↓
  Repair/add VBA header
       ↓
  Strip unnecessary attributes
       ↓
  Write as cp1252 (for Visio import)
       ↓
  Import via VBComponents.Import()
       ↓
  Convert back to UTF-8 (for editor)
```

### 5. File Watcher (`file_watcher.py`)

**Purpose:** Real-time bidirectional synchronization

**Features:**
- Monitors file system changes (using `watchdog`)
- Debounces rapid changes (avoids import loops)
- Polls Visio for changes (if bidirectional mode enabled)
- Handles file deletions with `--sync-delete-modules`

**Sync Strategies:**
- **Unidirectional (VS Code → Visio):** File watcher only
- **Bidirectional (VS Code ↔ Visio):** File watcher + polling

## Data Flow

### Export Flow
```
User runs: visiowings export --file doc.vsdm
       ↓
CLI initializes VisioVBAExporter
       ↓
Exporter connects to Visio via COM
       ↓
Document Manager discovers all open documents
       ↓
For each document with VBA:
   ├─ Calculate module content hash
   ├─ Compare with last export hash
   ├─ If changed: export modules to folder
   ├─ Strip headers, convert encoding
   ├─ Check for local file conflicts
   └─ Prompt user if conflicts exist
       ↓
Return exported files + hashes
```

### Import Flow
```
User saves file in VS Code
       ↓
File Watcher detects change
       ↓
Debounce timer (prevent rapid imports)
       ↓
Importer receives file path
       ↓
Repair VBA headers
       ↓
(Rubberduck: Inject @Folder annotations)
       ↓
Find target document (from folder structure)
       ↓
Check if module exists in Visio
   ├─ Document module? → Require --force flag
   ├─ Regular module? → Remove old, import new
   └─ New module? → Import directly
       ↓
Import via VBComponents.Import()
       ↓
Convert file back to UTF-8
```

### Bidirectional Sync Flow
```
┌────────────────────────────────────────────┐
│  File System (VS Code)                     │
│  *.bas, *.cls, *.frm files                 │
└──────────┬─────────────────────────────────┘
           │
           │  watchdog monitors changes
           │  (on_modified event)
           ▼
    ┌──────────────┐
    │ File Watcher │◄─────────┐
    └──────┬───────┘          │
           │                   │
           │ imports           │ polls every 4s
           │                   │
           ▼                   │
    ┌──────────────┐          │
    │ VBA Importer │          │
    └──────┬───────┘          │
           │                   │
           ▼                   │
    ┌──────────────┐   exports │
    │    Visio     │───────────┘
    │  VBProject   │
    └──────────────┘
```

## Threading Model

**Main Thread:**
- CLI initialization
- Initial export
- File watcher startup

**File Watcher Thread:**
- Monitors file system events
- Triggers imports on file changes
- Requires COM initialization (`pythoncom.CoInitialize()`)

**Polling Thread (Bidirectional only):**
- Periodically checks Visio for changes
- Triggers exports when hashes differ
- Requires separate COM initialization

**Thread Safety:**
- Each thread calling COM must initialize/uninitialize COM
- Debouncing prevents race conditions
- Hash comparison prevents export loops

## Error Handling Strategy

### Connection Errors
- Reconnect automatically on lost COM connection
- Graceful degradation if Visio is closed

### Import/Export Errors
- Skip problematic files, continue with others
- Log errors with `--debug` flag
- User-friendly error messages

### User Conflicts
- Interactive prompts for file conflicts
- Options: overwrite all, skip all, interactive, cancel
- Per-file granularity (not all-or-nothing)

## Performance Considerations

### Change Detection
- MD5 hashing avoids unnecessary exports
- Content normalization reduces false positives
- Debouncing reduces redundant imports

### File I/O
- Batch operations where possible
- Lazy loading of file contents
- Temp file cleanup

### Memory
- Streaming file reads/writes
- Limited caching (only hashes)
- Explicit COM object cleanup

## Security Considerations

### VBA Macro Trust
- Requires "Trust access to VBA project object model"
- User must explicitly enable in Visio settings

### File System
- No arbitrary file execution
- Limited to `.bas`, `.cls`, `.frm` extensions
- Path validation and sanitization

### COM Security
- Uses existing Visio instance (no remote COM)
- No credential storage
- Read/write access limited to VBA modules
