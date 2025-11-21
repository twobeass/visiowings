# CLI Module Documentation

**File:** `visiowings/cli.py`

## Overview

The CLI module provides the command-line interface for visiowings using Python's `argparse` library. It orchestrates all user interactions and coordinates between the core components.

## Commands

### `edit` - Interactive Editing Mode

Full bidirectional synchronization with live file watching.

**Usage:**
```bash
visiowings edit --file document.vsdm [options]
```

**Options:**
- `--file, -f` (required): Path to Visio file
- `--output, -o`: Export directory (default: current directory)
- `--force`: Allow overwriting document modules (ThisDocument.cls)
- `--bidirectional`: Enable Visio → VS Code sync (polling)
- `--debug`: Enable verbose logging
- `--sync-delete-modules`: Auto-delete Visio modules when local files deleted

**Workflow:**
1. Parse arguments
2. Initialize and connect VBA Exporter
3. Perform initial export of all modules
4. Initialize VBA Importer
5. Start File Watcher with bidirectional sync (if enabled)
6. Run until user interrupts (Ctrl+C)

**Implementation:**
```python
def cmd_edit(args):
    # Initial export
    exporter = VisioVBAExporter(visio_file, debug)
    all_exported, all_hashes = exporter.export_modules(output_dir)
    
    # Setup importer
    importer = VisioVBAImporter(visio_file, force, debug)
    
    # Start watcher
    watcher = VBAWatcher(
        output_dir, 
        importer, 
        exporter=exporter,
        bidirectional=bidirectional,
        sync_delete_modules=sync_delete
    )
    watcher.last_export_hashes = all_hashes
    watcher.start()  # Blocks until Ctrl+C
```

### `export` - One-Time Export

Export VBA modules from Visio to file system (no watching).

**Usage:**
```bash
visiowings export --file document.vsdm --output ./vba_modules
```

**Options:**
- `--file, -f` (required): Path to Visio file
- `--output, -o`: Export directory
- `--debug`: Enable verbose logging

**Workflow:**
1. Connect to Visio
2. Export all modules
3. Display summary
4. Exit

### `import` - One-Time Import

Import VBA modules from file system into Visio (no watching).

**Usage:**
```bash
visiowings import --file document.vsdm --input ./vba_modules --force
```

**Options:**
- `--file, -f` (required): Path to Visio file
- `--input, -i`: Import directory
- `--force`: Allow overwriting document modules
- `--debug`: Enable verbose logging

**Workflow:**
1. Connect to Visio
2. Scan input directory for `.bas`, `.cls`, `.frm` files
3. Import each file (respecting folder structure)
4. Display summary
5. Exit

## Multi-Document Support

The CLI automatically handles multiple open Visio documents:

**Folder Structure:**
```
vba_modules/
├── drawing/              # Main drawing document
│   ├── Module1.bas
│   └── ThisDocument.cls
├── mystencil/            # Stencil document
│   └── StencilCode.bas
└── mytemplate/           # Template document
    └── TemplateHelpers.cls
```

**File Count Display:**
- Single document: `✓ 5 modules exported`
- Multiple documents: `✓ 12 modules exported from 3 documents`

## Argument Parsing

### Parser Structure
```python
main_parser = argparse.ArgumentParser()
subparsers = main_parser.add_subparsers(dest='command')

edit_parser = subparsers.add_parser('edit')
export_parser = subparsers.add_parser('export')
import_parser = subparsers.add_parser('import')
```

### Validation
- File path validation (must exist for export/edit)
- Directory creation (output dirs created if missing)
- Visio file extension check (`.vsdm`, `.vsdx`, `.vssm`, `.vssx`, `.vstm`, `.vstx`)

## Error Handling

### File Not Found
```python
if not visio_file.exists():
    print(f"❌ File not found: {visio_file}")
    return
```

### Connection Failures
```python
if not exporter.connect_to_visio():
    return  # Error already printed by exporter
```

### No Modules Found
```python
if not all_exported:
    print("❌ No modules exported")
    return
```

## Debug Mode

When `--debug` is enabled:
- Verbose component initialization
- Hash values displayed for each document
- Detailed error traces
- Thread synchronization logging

**Example Output:**
```
[DEBUG] Debug mode enabled
[DEBUG] Exporting drawing...
[DEBUG] Hash calculated: a3f7b219... (3 modules)
[DEBUG] mystencil: Hash 9e2c4d81... (2 modules)
```

## Entry Point

**Package Entry Point (`setup.py`):**
```python
entry_points={
    'console_scripts': [
        'visiowings=visiowings.cli:main',
    ],
}
```

**Direct Execution:**
```python
if __name__ == '__main__':
    main()
```

## Best Practices

### Always Specify Output Directory
```bash
# Good
visiowings edit --file doc.vsdm --output ./vba

# Avoid (clutters current directory)
visiowings edit --file doc.vsdm
```

### Use Debug Mode for Troubleshooting
```bash
visiowings edit --file doc.vsdm --debug
```

### Enable Bidirectional Sync Carefully
```bash
# Only when you want Visio changes to overwrite local edits
visiowings edit --file doc.vsdm --bidirectional
```

## Future Enhancements

- [ ] Config file support (`.visiowingsrc`)
- [ ] Profile-based settings
- [ ] Batch mode for multiple files
- [ ] Interactive mode improvements
- [ ] Progress bars for large exports
