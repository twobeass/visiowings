# Manual Testing Procedures for Visiowings

This manual testing guide provides comprehensive, step-by-step instructions to verify the reliability and correctness of visiowings after robustness improvements. Each section lists prerequisites, actions, and expected outcomes. 

## Table of Contents
- Prerequisites
- Core Functional Tests
    - Exporting VBA Modules
    - Importing VBA Modules
    - Live Editing (File Watcher)
    - Bidirectional Sync
    - Multi-Document Handling
- Robustness & Edge Case Tests
    - File System Race Conditions
    - Encoding Edge Cases
    - Error Handling (Permissions, Connection Loss)
    - Shutdown Reliability
    - Document Module Protection
    - Folder/Module Naming

---

## Prerequisites
- Windows OS, Python 3.8+
- Microsoft Visio with VBA support (Trust access to VBA project object model enabled)
- Test Visio files (at least one `.vsdm` with VBA code)
- Visual Studio Code (or compatible text editor)
- Git Bash or similar terminal

---

## Core Functional Tests

### 1. Exporting VBA Modules
**Goal:** Ensure modules are exported to disk in correct format

#### Steps:
1. Open a `.vsdm` or `.vsdx` macro-enabled Visio document with VBA code.
2. In terminal:
   ```bash
   visiowings export --file test.vsdm --output ./test_vba_exports
   ```
3. Confirm output directory is created with subfolders for each open document (drawing, stencil, template, etc).
4. Verify that `.bas`, `.cls`, `.frm` files are present.
5. Open files in your text editor and check for:
    - `Attribute VB_Name = "..."` present
    - No extra VERSION/Begin/End/MultiUse headers
    - Encoding is UTF-8 (check in VS Code status bar)

**Expected:**
- All modules exported, headers are stripped, encoding is correct.

### 2. Importing VBA Modules
**Goal:** Ensure modules can be imported from disk to Visio, repairing headers as needed

#### Steps:
1. Save a `.bas`/`.cls` file with edited code (try changing a Sub/adding new lines).
2. In terminal:
   ```bash
   visiowings import --file test.vsdm --input ./test_vba_exports
   ```
3. In Visio, open VBA editor and verify that changes are reflected.

**Expected:**
- Modules update as expected.
- No import errors for valid code.
- For files missing `Attribute VB_Name`, header is added automatically.

### 3. Live Editing (File Watcher)
**Goal:** Verify live sync (file changes in editor sync to Visio automatically)

#### Steps:
1. In terminal:
   ```bash
   visiowings edit --file test.vsdm --output ./test_live_sync
   ```
2. Edit a module in VS Code, save the file.
3. Within 1-2s, check console for "Change detected" and "Imported" messages.
4. In Visio VBA editor, verify that code is updated automatically.

**Expected:**
- Changes detected instantly and synced to Visio.
- No errors/logging of failures.

### 4. Bidirectional Sync
**Goal:** Confirm changes made in Visio are reflected back to code files

#### Steps:
1. Start watcher in bidirectional mode:
   ```bash
   visiowings edit --file test.vsdm --output ./test_bidi_sync --bidirectional
   ```
2. In Visio VBA editor, change/add a procedure, save Visio file.
3. Wait up to 5 seconds.
4. Open code file in VS Code and check if updates appear.

**Expected:**
- Changes in Visio exported back to files within polling interval.

### 5. Multi-Document Handling
**Goal:** Verify exports/imports work for open drawings, stencils, templates simultaneously

#### Steps:
1. Open primary drawing file and a custom stencil/template (each with VBA) in Visio.
2. Start watcher and export as before.
3. Confirm files appear in separate subfolders (e.g. `drawing/`, `mystencil/`).
4. Edit files in each folder and check sync/import/export for all.

**Expected:**
- File association by document folder works for all types.
- No cross-document overwrites or missing modules.

---

## Robustness & Edge Case Tests

### 6. File System Race Conditions
**Goal:** Check that rapid/simultaneous edits are handled gracefully

#### Steps:
1. Edit and save the same module file multiple times quickly (<1s between saves).
2. Confirm "Debouncing" messages appear in debug mode and sync does not loop or error.

**Expected:**
- Rapid changes only trigger one sync per second.

### 7. Encoding Edge Cases
**Goal:** Ensure encoding loss detection and warnings work

#### Steps:
1. Insert characters outside cp1252 (e.g. emoji [translate:😀], CJK [translate:你好]) in a VBA file and attempt import.
2. Watch for explicit warning messages in terminal about incompatible characters.
3. Verify that imported code does not contain broken or replaced characters (or is skipped with a warning).

**Expected:**
- Warnings shown for incompatible characters. No silent data corruption.

### 8. Error Handling (Permissions, Connection Loss)
**Goal:** Validate user experience in error situations

#### Steps:
1. Attempt import/export while Visio is closed.
2. Remove write permission from output directory and try exporting.
3. Try importing an intentionally corrupt module file (invalid VBA syntax, missing headers).

**Expected:**
- Clear error messages printed.
- No crash, application continues running or exits cleanly.

### 9. Shutdown Reliability
**Goal:** Ensure Ctrl+C / SIGINT exits gracefully every time

#### Steps:
1. While watcher is running, press Ctrl+C at different moments (idle, during active import/export, etc).
2. Confirm "Shutting down gracefully" message and all threads stop cleanly.

**Expected:**
- Console returns to prompt, no lingering python processes.

### 10. Document Module Protection
**Goal:** Ensure document modules (ThisDocument.cls) are protected by `--force` option

#### Steps:
1. Remove or edit `ThisDocument.cls` in code files.
2. Try importing without `--force`:
   ```bash
   visiowings import --file test.vsdm --input ./test_vba_exports
   ```
3. Confirm import is skipped with a warning.
4. Now re-run with `--force` flag and confirm import proceeds.

**Expected:**
- Document modules only overwritten with explicit `--force` option. Skips silently otherwise.

### 11. Folder/Module Naming Edge Cases
**Goal:** Confirm correct mapping and handling of odd file/folder names

#### Steps:
1. Use document/module names with spaces, hyphens, mixed case, and special characters.
2. Confirm all such names are mapped correctly to valid folders/files.
3. Ensure import/export does not fail or misassign code modules.

**Expected:**
- All files/folders processed and associated correctly, regardless of name complexity.

---

## How to Report Issues
- Document any test that fails, with steps to reproduce and output/error messages.
- Open issues on [GitHub](https://github.com/twobeass/visiowings/issues) with full details.
- For encoding or data loss bugs, attach a minimal example file.

---

## Summary
Following this procedure ensures all core features and critical edge cases in visiowings are tested thoroughly after major robustness or bug fix releases.