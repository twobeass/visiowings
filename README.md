# visiowings

> VBA Editor for Microsoft Visio with VS Code integration - inspired by xlwings

🚀 Edit your Visio VBA code in VS Code (or any editor you like) with **live synchronization** back to Visio!

## Features

- ✅ **Line numbers** (which the VBA editor doesn't have!)
- 🌙 **Dark mode** for easier reading
- 🔄 **Live sync** - changes in VS Code automatically update Visio
- 👁️ **Git version control** while editing (not after!)
- ⌨️ **Modern editor features**: multi-cursor, auto-complete, search & replace
- 🚀 All your favorite keyboard shortcuts (Ctrl+/, Shift+Alt+Arrow, etc.)

## Why visiowings?

The Visio VBA editor lacks modern features that developers expect. **visiowings** brings the power of VS Code to Visio development:

```bash
# Instead of this painful workflow:
Visio VBA Editor → No line numbers → Limited editing → Manual version control

# You get this:
VS Code → Full editor features → Live sync to Visio → Automatic Git tracking
```

## Installation

### Prerequisites

- **Windows** (required for COM automation)
- **Python 3.8+**
- **Microsoft Visio** (any version with VBA support)

### Install from GitHub

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

### 1. Enable VBA Project Access in Visio

⚠️ **Important**: Before using visiowings, you must enable VBA project access:

1. Open Visio
2. **File** → **Options** → **Trust Center** → **Trust Center Settings**
3. **Macro Settings** → ☑ "Trust access to the VBA project object model"

### 2. Open your Visio file

Make sure your Visio file is:
- Saved as `.vsdm` (macro-enabled format)
- Currently open in Visio

### 3. Start editing with live sync

```bash
cd /path/to/your/project
visiowings edit --file "C:/path/to/your/document.vsdm"
```

This will:
1. Export all VBA modules to the current directory
2. Start watching for file changes
3. Auto-sync any changes back to Visio

### 4. Edit in VS Code

```bash
code .  # Open VS Code in current directory
```

Now edit your `.bas`, `.cls`, or `.frm` files. Every time you save (Ctrl+S), the changes are **instantly synchronized** to Visio!

## Usage

### Edit Mode (with live sync)

```bash
visiowings edit --file document.vsdm
visiowings edit --file document.vsdm --output ./vba_modules
```

### Export Only

Export VBA modules without watching for changes:

```bash
visiowings export --file document.vsdm --output ./vba_modules
```

### Import Only

Import VBA modules from files back into Visio:

```bash
visiowings import --file document.vsdm --input ./vba_modules
```

## Example Workflow

```bash
# 1. Open your Visio file in Visio
# 2. Navigate to your project folder
cd C:/Projects/MyVisioProject

# 3. Start visiowings
visiowings edit --file "MyDiagram.vsdm"

# Output:
# 📂 Visio-Datei: C:\Projects\MyVisioProject\MyDiagram.vsdm
# 📁 Export-Verzeichnis: C:\Projects\MyVisioProject
#
# === Exportiere VBA-Module ===
# ✓ Exportiert: Module1.bas
# ✓ Exportiert: ClassModule1.cls
#
# ✓ 2 Module exportiert
#
# === Starte Live-Synchronisation ===
# 👁️  Überwache Verzeichnis: C:\Projects\MyVisioProject
# 💾 Speichere Dateien in VS Code (Ctrl+S) um sie nach Visio zu synchronisieren
# ⏸️  Drücke Ctrl+C zum Beenden...

# 4. Edit Module1.bas in VS Code and save (Ctrl+S)
# Output:
# 📝 Änderung erkannt: Module1.bas
# ✓ Importiert: Module1.bas

# 5. Check Visio - your changes are already there!
```

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

1. **VBA** (Wine-HQ or similar)
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
- `.cls` - Class modules
- `.frm` - User forms

## Troubleshooting

### "Trust access to the VBA project object model"

If you get an error about VBA project access:

1. Open Visio
2. **File** → **Options** → **Trust Center** → **Trust Center Settings**
3. **Macro Settings** → ☑ "Trust access to the VBA project object model"
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

## Comparison with xlwings

**visiowings** is heavily inspired by [xlwings](https://www.xlwings.org/)'s VBA editing feature:

| Feature | xlwings (Excel) | visiowings (Visio) |
|---------|----------------|--------------------|
| Edit VBA in VS Code | ✅ | ✅ |
| Live sync | ✅ | ✅ |
| Export/Import | ✅ | ✅ |
| Python ↔ VBA calls | ✅ | ❌ (not yet) |
| UDFs | ✅ | N/A |

## Roadmap

- [ ] Add Python ↔ Visio integration (like xlwings `RunPython`)
- [ ] Support for Visio templates (`.vstm`)
- [ ] Standalone executable (no Python required)
- [ ] GUI version
- [ ] Auto-backup before import
- [ ] Diff viewer for changes

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
