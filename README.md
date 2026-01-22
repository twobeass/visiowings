# visiowings

VBA Editor for Microsoft Visio with VS Code integration – Edit Visio VBA modules in your favorite editor with live sync, Git support, and modern tooling. Inspired by xlwings.

---

## Table of Contents
- [Features](#features)
- [Why visiowings?](#why-visiowings)
- [Installation](#installation)
- [Quick Start](#quick-start)
- [Usage & Commands](#usage--commands)
- [Advanced Features](#advanced-features)
- [Integration](#integration)
- [Reference](#reference)
- [Troubleshooting](#troubleshooting)
- [Roadmap](#roadmap)
- [Contributing](#contributing)
- [License & Acknowledgments](#license--acknowledgments)

---

## Features
- **Live sync:** Edit in VS Code, changes update Visio instantly
- **Bidirectional sync:** (optional) Changes in Visio update VS Code
- **Multi-document support:** Sync several open files at once
- **Git version control:** See live diffs in VS Code
- **Automatic change detection:** Smart, hash-based, and loop-proof
- **Module deletion sync:** Remove modules in Visio when deleted in VS Code
- **Modern VS Code features:** Line numbers, dark mode, search, shortcuts
- **Safe import/export:** Robust comparison and header management prevents unwanted overwrites and import failures

## Why visiowings?
Visio VBA Editor lacks modern tooling. visiowings brings:
- Full-featured editor experience
- Fast, bidirectional sync
- Real Git integration
- No more manual copy-paste or limited VBA interface
- Safe cross-platform editing and importing of modules

---

## Installation
### Prerequisites
- Windows (required for COM automation)
- Python 3.8+
- Microsoft Visio (with VBA support)

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

---

## Quick Start
### Single Document
1. Enable VBA access in Visio: Options → Trust Center → Macro Settings → ☑ "Trust access to the VBA project object model"
2. Open your Visio file (e.g. `yourdoc.vsdm`)
3. Start editing:
   ```bash
   visiowings edit --file yourdoc.vsdm --bidirectional
   ```
4. Edit `.bas`, `.cls`, `.frm` files in VS Code, save to sync

### Multiple Documents
1. Open all relevant Visio documents (drawings, stencils, templates)
2. Run:
   ```bash
   visiowings edit --file your_main_document.vsdx --bidirectional --force
   ```
3. Edit VBA files in VS Code; changes sync to matching Visio file

---

## Usage & Commands

### Edit Mode
```bash
# VS Code → Visio only
visiowings edit --file yourdoc.vsdm

# Bidirectional sync
visiowings edit --file yourdoc.vsdm --bidirectional

# Overwrite document modules
visiowings edit --file yourdoc.vsdm --force

# Enable module deletion & debug mode
visiowings edit --file yourdoc.vsdm --sync-delete-modules --debug

# Custom output directory
visiowings edit --file yourdoc.vsdm --output ./vba_modules

# Enable Rubberduck folder structure integration
visiowings edit --file yourdoc.vsdm --rubberduck
```

### Export & Import
```bash
# Export VBA modules (one-time)
visiowings export --file yourdoc.vsdm --output ./vba_modules --rd

# Force export of .frx files
visiowings export --file yourdoc.vsdm --export-frx

# Import VBA modules
visiowings import --file yourdoc.vsdm --input ./vba_modules --force --rd
```

### Rubberduck Integration (New!)
Use Rubberduck-compatible folder annotations (`@Folder("Folder.Sub")`) to organize your VBA project structure automatically.

- **Export:** Reads `@Folder` annotations from VBA modules and exports files into the corresponding directory structure.
- **Import:** Detects file location in the directory structure and automatically injects or updates `@Folder` annotations in the imported code.
- **Sync:** Works seamlessly with `edit` mode to maintain structure while you work.

### Safe Export & Import (NEW)
- **Export:** Uses normalization and header-stripping to prevent accidental differences and overwriting. On conflict, user is prompted to overwrite, skip, or choose interactively.
- **Import:** Before importing, headers are repaired and encoding normalized. Comments and Option Explicit are preserved. Document modules are handled safely (force option required).

### Command Reference
| Option                 | Description                                                              |
|------------------------|--------------------------------------------------------------------------|
| `--file`, `-f`         | Visio file path (`.vsdm`, `.vsdx`, `.vssm`, `.vssx`, `.vstm`, `.vstx`)    |
| `--output`, `-o`       | Export directory (default: current directory)                            |
| `--input`, `-i`        | Import directory (for import command)                                    |
| `--force`              | Overwrite Document modules (ThisDocument.cls)                            |
| `--bidirectional`      | Enable bidirectional sync (Visio ↔ VS Code)                              |
| `--sync-delete-modules`| Automatically delete Visio modules when matching files are deleted        |
| `--rubberduck`, `--rd` | Enable Rubberduck @Folder annotation support for directory structure      |
| `--export-frx`         | Force export of .frx files even if .frm code has not changed              |
| `--debug`              | Verbose debug logging                                                    |

---

## Advanced Features
- **Multi-document support:** Each Visio file with VBA is exported/imported to a dedicated folder
- **Rubberduck Integration:** Map `@Folder("Parent.Child")` annotations to real directory structures automatically
- **Module deletion sync:** When enabled, `.bas`, `.cls`, `.frm` deletes in VS Code remove corresponding modules in Visio
- **Smart change detection:** Only syncs when content changes; polling interval is optimized
- **Intelligent Form Export:** .frx files are now only exported if the corresponding .frm code changes to prevent Git churn. Use `--export-frx` to force export.
- **Bidirectional sync:** Changes in VS Code or Visio keep both in sync with the selected polling interval (default: 4 seconds)
- **NEW: Safe Import/Export:** Content checks, normalization, encoding handling, and interactive user options ensure maximum safety

---

## Integration
### VS Code Setup
- Associate `.bas`, `.cls`, `.frm` files:
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
- Recommended Extensions:
  - VBA syntax highlighting (`@ext:vba`)
  - GitLens for Git history/blame

### Git Integration
- Initialize git:
  ```bash
  git init
  git add *.bas *.cls
  git commit -m "Initial VBA modules"
  ```
- Use VS Code’s Git tools for diffs and versioning

---

## Reference
### Supported File Types
- `.bas` - Standard VBA modules
- `.cls` - Class modules (incl. Document modules)
- `.frm` - User forms

### Example Project Structure
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

---

## Troubleshooting

### Trust Center
Visio → Options → Trust Center → Macro Settings → ☑ "Trust access to the VBA project object model"

### Visio document not open
- Always open the document in Visio before starting visiowings

### No syncing/constant exporting
- Ensure file watcher is running and files are saved
- Document must be open in Visio
- Use `--debug` to diagnose issues

### Unicode/Encoding
- Use UTF-8 in VS Code
- **NEW:** All import/export routines auto-convert and normalize encoding for safe editing

### Document module updates
- Use `--force` to update ThisDocument.cls modules
- **NEW:** Document modules are handled with extra safety and user confirmation

---

## Roadmap
- [ ] Python <-> Visio integration (`RunPython`)
- [ ] Configurable polling interval
- [ ] Standalone executable (no Python required)
- [ ] GUI version
- [ ] Auto-backup before import
- [ ] Diff viewer
- [ ] `.visiowingsignore` file support
- [ ] **Improve documentation with more advanced sync/import examples**

---

## Contributing

Help and PRs welcome!

---

## License & Acknowledgments
MIT License – see [LICENSE](LICENSE)

- Inspired by [xlwings](https://www.xlwings.org/) by Felix Zumstein
- Built with [pywin32](https://github.com/mhammond/pywin32) & [watchdog](https://github.com/gorakhargosh/watchdog)

---

**twobeass** – [GitHub](https://github.com/twobeass)

⭐ Star the project, if helpful!
