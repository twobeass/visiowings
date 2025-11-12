# visiowings

> VBA Editor for Microsoft Visio with VS Code integration - inspired by xlwings

🚀 Edit your Visio VBA code in VS Code (or any editor you like) with **live synchronization** back to Visio!

---

## ⚡ Important: Supported Visio File Types

| Format        | Macro support | Recommended for VBA |
|-------------- |-------------- |-------------------- |
| .vsdm         | ✅ Yes        | ✔️ (Drawing, macros) |
| .vssm         | ✅ Yes        | ✔️ (Stencil, macros) |
| .vstm         | ✅ Yes        | ✔️ (Template, macros) |
| .vsdx/.vssx/.vstx | ❌ No    | ✖️ (No VBA support)   |

> **Note:** Only files in the .vsdm, .vssm, or .vstm formats can actually store VBA macros. Always save Visio drawings as `.vsdm` if using VBA features! Do not use `.vsdx`, `.vssx`, `.vstx` for any macro-related workflow. See [Microsoft Docs][2][4]

---

## Security Warning: Macro Automation & Trust Center
⚠️  **Never leave 'Trust access to the VBA project object model' permanently enabled!** 
- Enable only for the duration of VBA synchronization.
- Consider using digitally signed macros and [Trusted Locations][9] for safer automation.
- See [MS Office Security Notes][5][6][59] for recommended practices.

```bash
# Instead of this painful workflow:
Visio VBA Editor -> No line numbers -> Limited editing -> Manual version control

# You get this:
VS Code -> Full editor features -> Live sync to Visio -> Automatic Git tracking
```

- UserForms consist of two files: `.frm` (text) _and_ `.frx` (resources, binary)
- Both must be under version control and present during import/export; missing .frx leads to broken forms, lost controls, or import errors ([3][8]).
- ⚠️ `.frx` files cannot be merged/diffed with Git and must be marked as binary (see below).

#### Supported File Types (Extended)
- `.bas` Standard VBA module
- `.cls` Class module
- `.frm + .frx` UserForm (always commit/import/export both)

---

## Recommended .gitattributes for VBA (Windows and Visio-safe)

```gitattributes
*.bas text eol=crlf working-tree-encoding=CP1252
*.cls text eol=crlf working-tree-encoding=CP1252
*.frm text eol=crlf working-tree-encoding=CP1252
*.frx binary
```

> This ensures all VBA text files use Windows line endings (CRLF) and ANSI encoding. UserForm .frx is always treated as binary. This prevents common import bugs due to LF endings or wrong encoding [7][10][63].

---

## Quick Start (macro-enabled formats only)

### 1. Enable VBA Project Access in Visio

⚠️  Enable only temporarily as needed!

### 2. Open your macro-enabled Visio file (.vsdm/.vssm/.vstm)

### 3. Start editing with live sync
```sh
visiowings edit --file "C:/path/to/your/document.vsdm"
```

### 4. Edit in VS Code
- Make sure both `.frm` & `.frx` are always present for UserForms.

### 5. Importing
- Always import/export UserForms as `.frm` + `.frx` pair
- Missing `.frx` will break the form or lose controls/resources
```bash
code .  # Open VS Code in current directory
```

Now edit your `.bas`, `.cls`, or `.frm` files. Every time you save (Ctrl+S), the changes are **instantly synchronized** to Visio!

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

# All options combined
visiowings edit --file document.vsdm --force --bidirectional --debug

# Custom output directory
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

```bash
# With debug mode, you can see the hash comparison:
visiowings edit --file document.vsdm --bidirectional --debug

# Output:
# [DEBUG] Hash berechnet: 882c423e... (3 Module)
# [DEBUG] Last hash: 882c423e...
# [DEBUG] Current hash: 882c423e...
# [DEBUG] Hashes identisch - kein Export
# [DEBUG] Keine Änderungen in Visio erkannt, kein Export.
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

---

## Troubleshooting
- If a UserForm (form) fails to render, check if the `.frx` file is missing for its `.frm`.
- Never use .vsdx/.vssx/.vstx for VBA macros ([2][4]).
- Multiple Visio processes open: Always select the right window/document and avoid duplicate file names in different Visio instances.
- Using `--force` on ThisDocument.cls overwrites ALL events & document logic. Always backup first!
- Polling interval is currently fixed at 4s; consider making it configurable for advanced use cases.

---

## Contribution, Compliance, and Release Best Practices
- Publish wheels and source releases to PyPI using `pyproject.toml`—allows `pip install visiowings`.
- Reference [Keep a Changelog](https://keepachangelog.com) and semantic versioning.
- Prefer digitally signed macros & Trusted Locations for secure team use.
- Document your team's encoding/git requirements in the project root (see .gitattributes above).
---

## Links & Sources
See bottom of README or PR description for detailed references (Microsoft, Stack Overflow, Chandoo, etc.).
