# Human UAT (User Acceptance Testing)

This document is a **release-readiness checklist** for visiowings. Run it
end-to-end before cutting a release, after major refactors, or whenever
you want signed-off confidence that everything still works on a real
machine.

The audience is a human tester sitting at a Windows box with Visio
installed. Each section is a self-contained scenario with:

- **Goal** — what we're proving
- **Setup** — what to prepare beforehand
- **Steps** — copy-pasteable actions
- **Expected** — what success looks like
- **Pass / Fail** — fill in as you go

> 💡 Read [Manual testing](manual-testing.md) first — it covers the same
> core features as a tutorial. **This page is the audit checklist** with
> stricter pass/fail criteria and explicit coverage of the Phase A–H
> tooling and supply-chain artefacts.

---

## Tester information

Fill in before starting:

| Field | Value |
| --- | --- |
| Tester name | |
| Date | |
| visiowings version under test | |
| Git commit / tag | |
| OS + version | |
| Visio version | |
| Python version | |
| VS Code version | |
| Test environment (laptop, VM, …) | |

---

## A. Environment prerequisites

> Run these first. If any fails, fix the environment before continuing —
> later sections assume A passes.

### A1. Trust Center allows VBA project access

| Step | Action |
| --- | --- |
| 1 | Open Visio. |
| 2 | File → Options → Trust Center → Trust Center Settings → Macro Settings. |
| 3 | Verify ☑ **Trust access to the VBA project object model** is checked. |

**Expected:** Checkbox is enabled. Without this, every COM-based
operation will fail with a generic permission error.

| ☐ Pass | ☐ Fail | Notes: |
| --- | --- | --- |

### A2. Test fixtures are ready

| Step | Action |
| --- | --- |
| 1 | Have at least one `.vsdm` file with at least one `.bas`, one `.cls` (incl. `ThisDocument`), and one `.frm` module with non-trivial code. |
| 2 | Optionally, have one `.vssm` stencil and one `.vstm` template open alongside (for multi-document tests in section F). |
| 3 | Note the file path(s) — copy them into the table below. |

| Fixture | Path |
| --- | --- |
| Main `.vsdm` | |
| Stencil (optional) | |
| Template (optional) | |

| ☐ Pass | ☐ Fail | Notes: |
| --- | --- | --- |

### A3. Python and pipx available

| Step | Command | Expected output |
| --- | --- | --- |
| 1 | `python --version` | `Python 3.10.x`, `3.11.x`, `3.12.x`, or `3.13.x` |
| 2 | `pipx --version` | Version printed (any) |

| ☐ Pass | ☐ Fail | Notes: |
| --- | --- | --- |

---

## B. Installation paths

Test each install method on a clean(-ish) system. If you're on the same
machine, uninstall (`pipx uninstall visiowings`) between B1, B2, B3.

### B1. `pipx install` (recommended path for end users)

| Step | Action |
| --- | --- |
| 1 | `pipx install visiowings` |
| 2 | `visiowings --version` |
| 3 | `visiowings --help` |

**Expected:** Install completes without errors. `--version` prints the
expected version. `--help` shows top-level commands (`edit`, `export`,
`import`, `init`).

| ☐ Pass | ☐ Fail | Notes: |
| --- | --- | --- |

### B2. `pip install` (alternative)

| Step | Action |
| --- | --- |
| 1 | `python -m venv .uat-venv && .uat-venv\Scripts\activate` |
| 2 | `pip install visiowings` |
| 3 | `visiowings --version` |

**Expected:** Same as B1.

| ☐ Pass | ☐ Fail | Notes: |
| --- | --- | --- |

### B3. Standalone EXE (no Python required)

| Step | Action |
| --- | --- |
| 1 | Download `visiowings.exe` from the latest GitHub Release. |
| 2 | Place it in a PATH directory (e.g. `C:\Tools\`). |
| 3 | Open a fresh terminal where Python is **not** on PATH. |
| 4 | `visiowings --version` |
| 5 | `visiowings --help` |

**Expected:** EXE runs, version matches the Release tag. No Python
installation needed. First launch may take 2–3 seconds (PyInstaller
bootstrap is normal).

| ☐ Pass | ☐ Fail | Notes: |
| --- | --- | --- |

### B4. Install from source (developer path)

| Step | Action |
| --- | --- |
| 1 | `git clone https://github.com/twobeass/visiowings.git` |
| 2 | `cd visiowings && just install` (or `pip install -e ".[dev,docs]" && pre-commit install`) |
| 3 | `just lint && just test` |
| 4 | `just info` |

**Expected:** All recipes succeed. `just lint` and `just test` exit
0; `just info` prints version + help.

| ☐ Pass | ☐ Fail | Notes: |
| --- | --- | --- |

---

## C. First-run experience and CLI surface

### C1. `visiowings init` wizard (Phase G)

| Step | Action |
| --- | --- |
| 1 | In an empty directory, run `visiowings init`. |
| 2 | Follow the prompts: pick a Visio file, output directory, codepage. |
| 3 | When done, `dir` (or `ls`) the directory. |
| 4 | Open the generated `.visiowings.toml` in an editor. |

**Expected:**
- Wizard discovers any open Visio documents and offers them as defaults.
- A `.visiowings.toml` file is written.
- The TOML contains keys for `file`, `output`, `codepage`, etc.
- Re-running `visiowings init` warns before overwriting.

| ☐ Pass | ☐ Fail | Notes: |
| --- | --- | --- |

### C2. Config layering — flag wins over `.visiowings.toml`

| Step | Action |
| --- | --- |
| 1 | In a directory with `.visiowings.toml` containing `output = "./vba"`, run `visiowings export --file <file>`. |
| 2 | Confirm export goes to `./vba`. |
| 3 | Now run `visiowings export --file <file> --output ./other`. |
| 4 | Confirm export goes to `./other`, **not** `./vba`. |

**Expected:** CLI flag overrides config; config overrides defaults.

| ☐ Pass | ☐ Fail | Notes: |
| --- | --- | --- |

### C3. CLI input validation (Phase E)

| Step | Action | Expected |
| --- | --- | --- |
| 1 | `visiowings export --file does_not_exist.vsdm` | Clean error: file not found, no traceback. |
| 2 | `visiowings export --file <doc>.vsdm --codepage zzz123` | Error: unknown codepage. |
| 3 | `visiowings export --file <doc>.txt` | Error: unsupported extension. |
| 4 | `visiowings export --file <doc>.vsdm --output /readonly` (read-only path) | Error: cannot write to output dir. |

**Expected:** All four cases error with a *one-line* user-facing message
and exit code ≠ 0. No Python traceback.

| ☐ Pass | ☐ Fail | Notes: |
| --- | --- | --- |

### C4. `--debug` produces structured logs (Phase D)

| Step | Action |
| --- | --- |
| 1 | `visiowings export --file <doc>.vsdm --debug` |
| 2 | Observe terminal output. |

**Expected:** `[DEBUG]` lines visible. Lines are timestamped (or at least
prefixed) and include module names. No raw `print()` debug noise mixed
in.

| ☐ Pass | ☐ Fail | Notes: |
| --- | --- | --- |

---

## D. Core features (smoke)

> These mirror [Manual testing](manual-testing.md) §1–§5. Run them
> quickly here as a release smoke; refer to the manual-testing doc for
> deeper exploration.

### D1. Export

| Step | Action |
| --- | --- |
| 1 | `visiowings export --file <doc>.vsdm --output ./uat-export` |
| 2 | Inspect `./uat-export/`. |

**Expected:**
- One subfolder per open Visio doc with VBA.
- `.bas`, `.cls`, `.frm` files inside.
- `Attribute VB_Name = "..."` retained, but VERSION/Begin/End/MultiUse
  headers stripped.
- Files are UTF-8 with line endings as configured (LF by default; check
  the `.editorconfig` association — `.bas/.cls/.frm` should be CRLF).

| ☐ Pass | ☐ Fail | Notes: |
| --- | --- | --- |

### D2. Import (round-trip)

| Step | Action |
| --- | --- |
| 1 | Edit one `.bas` file from D1; add a comment. |
| 2 | `visiowings import --file <doc>.vsdm --input ./uat-export` |
| 3 | Open Visio's VBA Editor (Alt+F11) and inspect the module. |

**Expected:** Module shows the new comment. No `Attribute` line is
duplicated. No encoding artefacts.

| ☐ Pass | ☐ Fail | Notes: |
| --- | --- | --- |

### D3. Live edit mode

| Step | Action |
| --- | --- |
| 1 | `visiowings edit --file <doc>.vsdm --output ./uat-edit` |
| 2 | In VS Code, edit a `.bas` file from `./uat-edit/`, save. |
| 3 | Within ~2s, observe terminal logs. |
| 4 | In Visio's VBA Editor, refresh and check the change is there. |

**Expected:** "Change detected" + "Imported" within 2s. Module updated
in Visio.

| ☐ Pass | ☐ Fail | Notes: |
| --- | --- | --- |

### D4. Bidirectional sync

| Step | Action |
| --- | --- |
| 1 | `visiowings edit --file <doc>.vsdm --output ./uat-bidi --bidirectional` |
| 2 | In Visio's VBA Editor, add a new function and save the document. |
| 3 | Wait up to 5s. |
| 4 | Open the corresponding `.bas` in VS Code. |

**Expected:** New function appears in the local file within the
configured polling interval (default 4s).

| ☐ Pass | ☐ Fail | Notes: |
| --- | --- | --- |

### D5. Module deletion sync

| Step | Action |
| --- | --- |
| 1 | `visiowings edit --file <doc>.vsdm --output ./uat-del --sync-delete-modules` |
| 2 | In VS Code, delete a `.bas` file (not `ThisDocument.cls`). |
| 3 | Open Visio VBA Editor. |

**Expected:** The module is gone in Visio. Console reports the removal.
Document modules (`ThisDocument`) are **never** deleted by this flag.

| ☐ Pass | ☐ Fail | Notes: |
| --- | --- | --- |

---

## E. Encoding and codepages (Phase E)

### E1. German + cp1252 round-trip

| Step | Action |
| --- | --- |
| 1 | In a `.bas` file, insert a comment with German umlauts: `' Größe für Zähler übergeben`. |
| 2 | Run `visiowings import --file <doc>.vsdm --input ...` (Visio document is German LCID). |
| 3 | Open Visio VBA Editor. |

**Expected:** Umlauts render correctly (no `Gr??e`). Reload of the
exported file in VS Code also shows correct characters.

| ☐ Pass | ☐ Fail | Notes: |
| --- | --- | --- |

### E2. Codepage override

| Step | Action |
| --- | --- |
| 1 | `visiowings export --file <doc>.vsdm --codepage cp1251` (force Cyrillic) |
| 2 | Re-import. |

**Expected:** No silent corruption. Either succeeds or warns explicitly
that the override conflicts with the document's LCID.

| ☐ Pass | ☐ Fail | Notes: |
| --- | --- | --- |

### E3. BOM detection

| Step | Action |
| --- | --- |
| 1 | Save a `.bas` file from VS Code with **UTF-8 with BOM** encoding. |
| 2 | Run `visiowings import --file <doc>.vsdm --input ...`. |

**Expected:** Import succeeds, BOM is stripped before Visio sees it (no
`` character at the start of the first line in the VBA Editor).

| ☐ Pass | ☐ Fail | Notes: |
| --- | --- | --- |

### E4. Out-of-range character warning

| Step | Action |
| --- | --- |
| 1 | Insert an emoji 😀 or CJK character 你好 in a `.bas` comment. |
| 2 | Run import against a cp1252 document. |

**Expected:** Tool prints an explicit warning naming the offending
character and either skips the file or replaces with `?` and continues.
**No silent data loss.**

| ☐ Pass | ☐ Fail | Notes: |
| --- | --- | --- |

---

## F. Multi-document and Rubberduck

### F1. Drawing + stencil + template open simultaneously

| Step | Action |
| --- | --- |
| 1 | Open all three fixture files in Visio. |
| 2 | `visiowings export --file <main>.vsdm --output ./uat-multi` |
| 3 | List `./uat-multi/`. |

**Expected:** Three subfolders, one per document, named after the
document (sanitised — no spaces or special chars). Modules from each
document only land in *their* folder.

| ☐ Pass | ☐ Fail | Notes: |
| --- | --- | --- |

### F2. Rubberduck `@Folder` export

| Step | Action |
| --- | --- |
| 1 | In a module, add as the first non-attribute line: `'@Folder("UI.Buttons")`. |
| 2 | `visiowings export --file <doc>.vsdm --output ./uat-rd --rd` |

**Expected:** File exported to `./uat-rd/<doc>/UI/Buttons/<Module>.bas`.

| ☐ Pass | ☐ Fail | Notes: |
| --- | --- | --- |

### F3. Rubberduck `@Folder` import (auto-inject)

| Step | Action |
| --- | --- |
| 1 | Move a `.bas` file from `./uat-rd/<doc>/` to `./uat-rd/<doc>/Helpers/`. |
| 2 | `visiowings import --file <doc>.vsdm --input ./uat-rd --rd` |
| 3 | Open the module in Visio. |

**Expected:** Code now contains `'@Folder("Helpers")` near the top.

| ☐ Pass | ☐ Fail | Notes: |
| --- | --- | --- |

---

## G. Resilience and error scenarios (Phase E)

### G1. Visio not running

| Step | Action | Expected |
| --- | --- | --- |
| 1 | Close Visio entirely. | — |
| 2 | `visiowings export --file <doc>.vsdm --output ./uat-noex` | Clean error: "Document not found / Visio not running"; no traceback; exit ≠ 0. |

| ☐ Pass | ☐ Fail | Notes: |
| --- | --- | --- |

### G2. Visio crashes mid-watch (COM reconnection)

| Step | Action |
| --- | --- |
| 1 | Start `visiowings edit --file <doc>.vsdm --bidirectional`. |
| 2 | Force-kill Visio via Task Manager. |
| 3 | Watch the terminal for ~10s. |
| 4 | Re-open the same document in Visio. |
| 5 | Save a file in VS Code. |

**Expected:** Up to 3 reconnect attempts (Phase D `_retry`). On final
failure, a clear `COMConnectionError` line. After Visio is back, sync
either auto-recovers or requires a restart with a clear "please restart"
message — **never** silent failure.

| ☐ Pass | ☐ Fail | Notes: |
| --- | --- | --- |

### G3. Document module protection

| Step | Action |
| --- | --- |
| 1 | Edit `ThisDocument.cls` locally (add a Sub). |
| 2 | `visiowings import --file <doc>.vsdm --input ...` (no `--force`). |
| 3 | Confirm warning printed and module not changed. |
| 4 | Re-run with `--force`. |

**Expected:** Step 3 skips with explicit warning. Step 4 imports.

| ☐ Pass | ☐ Fail | Notes: |
| --- | --- | --- |

### G4. Read-only output directory

| Step | Action |
| --- | --- |
| 1 | Make `./uat-readonly` read-only at the filesystem level. |
| 2 | `visiowings export --file <doc>.vsdm --output ./uat-readonly` |

**Expected:** Clear "permission denied" error referencing the path. Exit
≠ 0. No partial writes left over.

| ☐ Pass | ☐ Fail | Notes: |
| --- | --- | --- |

### G5. Graceful Ctrl+C

| Step | Action |
| --- | --- |
| 1 | Start `visiowings edit ... --bidirectional`. |
| 2 | Press Ctrl+C. |
| 3 | Repeat at three different moments: idle, mid-import, mid-export. |
| 4 | After exit, check Task Manager for stray `python.exe` / `visiowings.exe` processes. |

**Expected:** Each Ctrl+C produces "Shutting down…" within ~2s. No
zombie processes.

| ☐ Pass | ☐ Fail | Notes: |
| --- | --- | --- |

### G6. Rapid-fire saves are debounced

| Step | Action |
| --- | --- |
| 1 | Start `visiowings edit ... --debug`. |
| 2 | Save the same file 5 times within 1 second (Ctrl+S spam). |

**Expected:** "Debouncing" lines visible; only one import is performed
per debounce window (1s). No infinite loop.

| ☐ Pass | ☐ Fail | Notes: |
| --- | --- | --- |

---

## H. PyPI update check (Phase G)

### H1. Update available, opt-in default

| Step | Action |
| --- | --- |
| 1 | Install a deliberately old version: `pipx install visiowings==<old>` |
| 2 | Run any command (e.g. `visiowings --help`). |

**Expected:** A non-fatal "A newer version (X.Y.Z) is available" hint is
printed. Hint appears at most once per day (cached).

| ☐ Pass | ☐ Fail | Notes: |
| --- | --- | --- |

### H2. Opt-out via env var

| Step | Action |
| --- | --- |
| 1 | `set VISIOWINGS_NO_UPDATE_CHECK=1` (PowerShell: `$env:VISIOWINGS_NO_UPDATE_CHECK="1"`) |
| 2 | Run any command. |

**Expected:** No update-check network call (verify with offline test
below or with debug logging). No "newer version" hint.

| ☐ Pass | ☐ Fail | Notes: |
| --- | --- | --- |

### H3. Opt-out via `.visiowings.toml`

| Step | Action |
| --- | --- |
| 1 | In `.visiowings.toml`, set `update_check = false`. |
| 2 | Unset the env var, run any command. |

**Expected:** No update hint.

| ☐ Pass | ☐ Fail | Notes: |
| --- | --- | --- |

### H4. Offline behaviour

| Step | Action |
| --- | --- |
| 1 | Disconnect network (or block PyPI in firewall). |
| 2 | Run any command. |

**Expected:** No traceback, no >5s hang. Update check fails silently
(only visible in `--debug`).

| ☐ Pass | ☐ Fail | Notes: |
| --- | --- | --- |

---

## I. Documentation and discoverability

### I1. Docs site is reachable and current

| Step | Action |
| --- | --- |
| 1 | Open <https://twobeass.github.io/visiowings/>. |
| 2 | Verify the version in the footer matches the version under test. |
| 3 | Click through Getting Started → Configuration → Codepages, and Contributing → Development environment, Releasing, Optimization plan. |

**Expected:** No 404s. Latest changelog is visible. Edit links go to
GitHub.

| ☐ Pass | ☐ Fail | Notes: |
| --- | --- | --- |

### I2. PyPI page is correct

| Step | Action |
| --- | --- |
| 1 | Open <https://pypi.org/project/visiowings/>. |
| 2 | Check version, README rendering (no broken images), license badge. |

**Expected:** README looks the same as the GitHub one (Markdown
rendering doesn't drop badges or sections). License shows MIT.

| ☐ Pass | ☐ Fail | Notes: |
| --- | --- | --- |

### I3. README badges are live

| Step | Action |
| --- | --- |
| 1 | Open <https://github.com/twobeass/visiowings>. |
| 2 | Check each badge resolves (CI, PyPI, Python, License, OpenSSF Scorecard, Docs). |

**Expected:** No "image not found", no stuck "unknown" badges.

| ☐ Pass | ☐ Fail | Notes: |
| --- | --- | --- |

---

## J. Release artefacts and supply-chain (Phase F + H)

### J1. Wheel + sdist install correctly

| Step | Action |
| --- | --- |
| 1 | Download `visiowings-X.Y.Z-py3-none-any.whl` and `visiowings-X.Y.Z.tar.gz` from the Release. |
| 2 | `pip install ./visiowings-X.Y.Z-py3-none-any.whl` in a fresh venv. |
| 3 | `visiowings --version`. |
| 4 | Repeat with the sdist. |

**Expected:** Both install cleanly. Version matches.

| ☐ Pass | ☐ Fail | Notes: |
| --- | --- | --- |

### J2. SBOM and license report present

| Step | Action |
| --- | --- |
| 1 | On the Release page, confirm `sbom.cdx.json` and `licenses.json` are attached. |
| 2 | Open `sbom.cdx.json` — check it's valid JSON and the top-level `bomFormat: CycloneDX`. |
| 3 | Open `licenses.json` — check it's a JSON array with one entry per dep. |

**Expected:** Both files present and well-formed.

| ☐ Pass | ☐ Fail | Notes: |
| --- | --- | --- |

### J3. Sigstore signature verification (Phase H)

| Step | Action |
| --- | --- |
| 1 | `pip install sigstore` (one-time). |
| 2 | Download `sbom.cdx.json` and `sbom.cdx.json.sigstore` to the same folder. |
| 3 | Run: |

```bash
sigstore verify github \
    --bundle sbom.cdx.json.sigstore \
    --cert-identity 'https://github.com/twobeass/visiowings/.github/workflows/publish.yml@refs/tags/vX.Y.Z' \
    --cert-oidc-issuer 'https://token.actions.githubusercontent.com' \
    sbom.cdx.json
```

**Expected:** `OK: sbom.cdx.json`. Repeat with `licenses.json` /
`licenses.json.sigstore`.

| ☐ Pass | ☐ Fail | Notes: |
| --- | --- | --- |

### J4. PyPI Trusted Publishing succeeded (no token leak)

| Step | Action |
| --- | --- |
| 1 | Open the `publish.yml` workflow run for the release tag. |
| 2 | Inspect the **PyPI upload** step output. |
| 3 | Confirm Repository **Settings → Secrets and variables → Actions** does **not** contain a `PYPI_API_TOKEN` secret. |

**Expected:** Upload uses `id-token: write` and OIDC. No long-lived
token configured anywhere.

| ☐ Pass | ☐ Fail | Notes: |
| --- | --- | --- |

### J5. OpenSSF Scorecard score did not regress (Phase H)

| Step | Action |
| --- | --- |
| 1 | Open <https://securityscorecards.dev/viewer/?uri=github.com/twobeass/visiowings>. |
| 2 | Note the overall score and compare to the previous release. |

**Expected:** Score is ≥ previous release. Any individual check that
*dropped* is intentional or has an open issue.

Previous score: ____ Current score: ____

| ☐ Pass | ☐ Fail | Notes: |
| --- | --- | --- |

---

## K. CI / automation sanity (informational)

These confirm the dev pipeline still works — they're not strictly UAT,
but worth a glance pre-release.

### K1. Latest `main` is green

| Step | Action |
| --- | --- |
| 1 | Open <https://github.com/twobeass/visiowings/actions>. |
| 2 | Latest run on `main` is all green. |

| ☐ Pass | ☐ Fail | Notes: |
| --- | --- | --- |

### K2. release-please PR exists (if there are unreleased commits)

| Step | Action |
| --- | --- |
| 1 | Open Pull Requests. |
| 2 | Look for an open `chore(main): release X.Y.Z` PR. |

**Expected:** One open release PR if and only if `main` has unreleased
`feat:` / `fix:` commits since the last tag.

| ☐ Pass | ☐ Fail | Notes: |
| --- | --- | --- |

### K3. Dependabot is healthy

| Step | Action |
| --- | --- |
| 1 | Open Pull Requests filtered by `author:app/dependabot`. |
| 2 | Review any open PRs. |

**Expected:** No PR has been open longer than ~14 days without a label
or response. Stale ones indicate a process gap, not a release blocker.

| ☐ Pass | ☐ Fail | Notes: |
| --- | --- | --- |

---

## L. Sign-off

> Only sign off after every section above is **Pass** (or has an
> attached, accepted exception).

| Field | Value |
| --- | --- |
| Total sections | 33 |
| Pass | |
| Fail | |
| Skipped (with reason) | |
| Tester signature | |
| Date | |
| Approved for release? | ☐ Yes ☐ No |

If **No**: open a GitHub issue per failed section using the
[Bug report template](https://github.com/twobeass/visiowings/issues/new?template=bug.yml),
attach this checklist, and block the release until resolved.

---

## How to report a UAT failure

For every failed step:

1. Note the section ID (e.g. `G2`).
2. Capture the exact CLI invocation and the full terminal output.
3. Capture the OS/Visio/Python versions from the *Tester information*
   table.
4. Open a GitHub issue with title `UAT failure: <section ID> – <short
   description>`.
5. For encoding or data-loss bugs, attach a *minimal reproducer* —
   smallest possible `.vsdm` + `.bas` that reproduces the issue.
6. Mark the issue as `release-blocker` if it would prevent shipping.
