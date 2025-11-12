# visiowings

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.8+](https://img.shields.io/badge/python-3.8%2B-blue.svg)](https://www.python.org/downloads/)

> VBA Editor for Microsoft Visio with VS Code integration - inspired by xlwings

🚀 Edit your Visio VBA code in VS Code (or any editor you like) with **live synchronization** back to Visio!

## Table of Contents
- [Features](#features)
- [Installation](#installation)
- [Quick Start](#quick-start)
- [Usage](#usage)
- [Troubleshooting](#troubleshooting)
- [Contributing](#contributing)

## Features

- ✅ **Line numbers** (which the VBA editor doesn't have!)
- 🌙 **Dark mode** for easier reading
- 🔄 **Live sync** - changes in VS Code automatically update Visio
- 🔄 **Bidirectional sync** - changes in Visio automatically update VS Code (optional)
- 👁️ **Git version control** while editing (not after!)
- ⏱️ **Smart change detection** - only syncs when content actually changed
- ⌨️ **Modern editor features**: multi-cursor, auto-complete, search & replace
- 🚀 All your favorite keyboard shortcuts (Ctrl+/, Shift+Alt+Arrow, etc.)
- 🐛 **Debug mode** for troubleshooting

## Multi-Document Support

- visiowings can sync VBA code from multiple open Visio files: drawings (`.vsdx`, `.vsdm`), stencils (`.vssm`, `.vssx`), and templates (`.vstm`, `.vstx`).
- Each Visio file with VBA code is exported to its own subfolder under your chosen project directory.

### Example Structure

```text
my_project/
├── drawing.vsdx
├── mystencil.vssm
└── vba_modules/
    ├── drawing/
    │   └── Module1.bas
    └── mystencil/
        ├── ThisDocument.cls
        └── StencilHelpers.bas
```

### How It Works

- **Automatic detection:** Any open Visio document with VBA code will be picked up.
- **Seamless import/export:** When you save code in VS Code, changes go instantly back to the correct Visio file. If you edit VBA in Visio, changes will sync back to the right file in VS Code on the next polling interval.
- **Subfolder matching:** During import, files in `vba_modules/drawing/` are synced back to the matching open Visio file named “drawing”.

### Quick Start for Multiple Documents

1. Open all relevant Visio documents (drawings, stencils, templates).
2. Run:
   ```bash
   visiowings edit --file your_main_document.vsdx --bidirectional --force
   ```
3. Edit your VBA files in VS Code. Changes always sync back to their matching Visio document.

### Notes

- If only one document with VBA code is open, all modules appear in a single subfolder.
- File/folder names are sanitized to avoid issues with special characters or spaces.

## Why visiowings?

The Visio VBA editor lacks modern features that developers expect. **visiowings** brings the power of VS Code to Visio development:

```text
# Instead of this painful workflow:
Visio VBA Editor -> No line numbers -> Limited editing -> Manual version control

# You get this:
VS Code -> Full editor features -> Live sync to Visio -> Automatic Git tracking
```

## Installation

### Prerequisites

- **Windows** (required for COM automation)
- **Python 3.8+**
- **Microsoft Visio** (any version with VBA support)


### Install from GitHub (recommended)

```bash
pip install git+https://github.com/twobeass/visiowings.git
```

### Install from source

```bash
git clone https://github.com/twobeass/visiowings.git
cd visiowings
pip install -e .
```

## Quick Start

1. **Enable VBA access**: Visio → File → Options → Trust Center → Trust Center Settings → Macro Settings → ☑ "Trust access to the VBA project object model"

2. **Install**: `pip install git+https://github.com/twobeass/visiowings.git`

3. **Open your Visio file** (saved as `.vsdm`)

4. **Start syncing**:
   ```bash
   visiowings edit --file "document.vsdm" --bidirectional
   code .
   ```

Done! Edit in VS Code, save with Ctrl+S, and changes sync instantly to Visio.

## Usage

### Edit Mode (with live sync)

```bash
# Basic mode (VS Code -> Visio only)
visiowings edit --file document.vsdm

# With bidirectional sync (VS Code <-> Visio)
visiowings edit --file document.vsdm --bidirectional

# Force overwrite Document modules (ThisDocument.cls)
visiowings edit --file document.vsdm --force

# Debug mode for troubleshooting
visiowings edit --file document.vsdm --debug

# Automatically delete Visio modules when .bas/.cls/.frm files are deleted locally
visiowings edit --file document.vsdm --sync-delete-modules

# All options combined
visiowings edit --file document.vsdm --force --bidirectional --sync-delete-modules --debug

# Custom output directory
visiowings edit --file document.vsdm --output ./vba_modules
```

### New Option: Module Deletion Sync

- `--sync-delete-modules`: When enabled, modules are automatically removed from Visio when their corresponding .bas/.cls/.frm files are deleted locally in VS Code.
- Default is **off**; activate explicitly if you want this behavior.

---

## Command Line Options

### `edit` command

| Option                  | Description                                                                 |
|------------------------|-----------------------------------------------------------------------------|
| `--file`, `-f`         | Visio file path (`.vsdm`) - **required**                                     |
| `--output`, `-o`       | Export directory (default: current directory)                                |
| `--force`              | Force overwrite Document modules (ThisDocument.cls)                          |
| `--bidirectional`      | Enable bidirectional sync (Visio <-> VS Code)                               |
| `--debug`              | Enable verbose debug logging                                                |
| `--sync-delete-modules`| Automatically delete Visio modules when local .bas/.cls/.frm files are deleted|

---

## Features
- 🧹 **Automatic module deletion**: If `--sync-delete-modules` is enabled, local file deletes remove the corresponding VBA module from Visio to maintain consistency.

## Example

```bash
# Enable automatic module delete
visiowings edit --file MyDiagram.vsdm --sync-delete-modules
```


### Export Only

Export VBA modules without watching for changes:

```bash
visiowings export --file document.vsdm --output ./vba_modules
```

### Import Only

Import VBA modules from files back into Visio:

```bash
visiowings import --file document.vsdm --input ./vba_modules --force
```

## Command Line Options

### `edit` command

| Option | Description |
|--------|-------------|
| `--file`, `-f` | Visio file path (`.vsdm`) - **required** |
| `--output`, `-o` | Export directory (default: current directory) |
| `--force` | Force overwrite Document modules (ThisDocument.cls) |
| `--bidirectional` | Enable bidirectional sync (Visio <-> VS Code) |
| `--debug` | Enable verbose debug logging |

### `export` command

| Option | Description |
|--------|-------------|
| `--file`, `-f` | Visio file path (`.vsdm`) - **required** |
| `--output`, `-o` | Export directory (default: current directory) |
| `--debug` | Enable verbose debug logging |

### `import` command

| Option | Description |
|--------|-------------|
| `--file`, `-f` | Visio file path (`.vsdm`) - **required** |
| `--input`, `-i` | Import directory (default: current directory) |
| `--force` | Force overwrite Document modules (ThisDocument.cls) |
| `--debug` | Enable verbose debug logging |

## Example Workflow

```bash
# 1. Open your Visio file in Visio
# 2. Navigate to your project folder
cd C:/Projects/MyVisioProject

# 3. Start visiowings with bidirectional sync
visiowings edit --file "MyDiagram.vsdm" --force --bidirectional

# Output:
# 📂 Visio file: C:\Projects\MyVisioProject\MyDiagram.vsdm
# 📁 Export directory: C:\Projects\MyVisioProject
#
# === Exporting VBA Modules ===
# ✓ Exported: ThisDocument.cls
# ✓ Exported: Module1.bas
# ✓ Exported: ClassModule1.cls
#
# ✓ 3 modules exported
#
# === Starting Live Synchronization ===
# 👁️  Watching directory: C:\Projects\MyVisioProject
# 💾 Save files in VS Code (Ctrl+S) to synchronize them to Visio
# 🔄 Bidirectional sync: Changes in Visio are automatically exported to VS Code.
# ⏸️  Press Ctrl+C to stop...

# 4. Edit Module1.bas in VS Code and save (Ctrl+S)
# Output:
# 📝 Change detected: Module1.bas
# ✓ Imported: Module1.bas

# 5. Edit VBA code in Visio (Alt+F11)
# Output (after ~4 seconds):
# 🔄 Visio document synchronized -> VS Code.

# 6. Check VS Code - your changes from Visio are already there!

```

## Bidirectional Sync

With the `--bidirectional` flag, visiowings enables two-way synchronization:

- **VS Code -> Visio**: Changes saved in VS Code (Ctrl+S) are immediately imported to Visio
- **Visio -> VS Code**: Changes in Visio VBA Editor are automatically exported to VS Code every 4 seconds

### Smart Change Detection

visiowings uses MD5 hash-based change detection to prevent unnecessary exports:

- Only exports when VBA code actually changes
- Prevents endless loops
- Pauses file watcher during export operations
- Efficient polling without constant file writes

## Git Integration

One of the **biggest benefits** is real-time Git integration:

```bash
# Initialize git in your project folder
git init
git add *.bas *.cls
git commit -m "Initial VBA modules"

# Now edit your VBA in VS Code
# Git will show you changes in real-time!
# Use VS Code's Git features:
# - See diffs immediately
# - Jump between changes (Alt+F5)
# - Stage/unstage specific changes
# - Commit with proper messages
```

## VS Code Setup

### Recommended Extensions

For the best experience, install these VS Code extensions:

1. **VBA syntax**
   - Provides syntax highlighting for `.bas`, `.cls`, `.frm` files
   - Search in VS Code: `@ext:vba`
2. **GitLens**
   - Enhanced Git integration
   - Inline blame and history

### Example VS Code Settings

```json
{
  "files.associations": {
    "*.bas": "vba",
    "*.cls": "vba",
    "*.frm": "vba"
  },
  "editor.tabSize": 4,
  "editor.insertSpaces": true
}
```

## Supported File Types

- `.bas` - Standard VBA modules
- `.cls` - Class modules (including Document modules like ThisDocument)
- `.frm` - User forms

## Troubleshooting

### "Trust access to the VBA project object model"

If you get an error about VBA project access:

1. Open Visio
2. **File** -> **Options** -> **Trust Center** -> **Trust Center Settings**
3. **Macro Settings** -> ☑ "Trust access to the VBA project object model"
4. Restart visiowings

### "Document not open"

The Visio document must be open in Visio before running visiowings commands.

### Unicode/Encoding Issues

If you have special characters (äöü), make sure your editor uses UTF-8 encoding.

### Changes not syncing

Make sure:
- The file watcher is running (you should see the 👁️ message)
- You're saving the file (Ctrl+S)
- The file extension is `.bas`, `.cls`, or `.frm`
- The document is still open in Visio

### Endless Loop / Constant Exports

This should not happen with the latest version, but if it does:

1. Use `--debug` flag to see what's triggering exports
2. Check the hash values in debug output
3. The file watcher is paused during exports to prevent triggering itself

### Document Module (ThisDocument.cls) not updating

Document modules require the `--force` flag:

```bash
visiowings edit --file document.vsdm --force
```

## Roadmap

- [ ] Add Python <-> Visio integration (like xlwings `RunPython`)
- [ ] Configurable polling interval
- [ ] Standalone executable (no Python required)
- [ ] GUI version
- [ ] Auto-backup before import
- [ ] Diff viewer for changes
- [ ] `.visiowingsignore` file support

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

MIT License - see [LICENSE](LICENSE) file for details.

## Acknowledgments

- Inspired by [xlwings](https://www.xlwings.org/) by Felix Zumstein
- Built with [pywin32](https://github.com/mhammond/pywin32) and [watchdog](https://github.com/gorakhargosh/watchdog)

## Author

**twobeass** - [GitHub](https://github.com/twobeass)

---

⭐ If you find this useful, please give it a star on GitHub!
