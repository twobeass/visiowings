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

---

## UserForms: Always Handle .frm and .frx Together

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
