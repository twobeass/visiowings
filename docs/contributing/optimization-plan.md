# The optimization & automation plan (Phases A–H)

Between January and May 2026 the project went through a structured eight-phase
modernisation that took it from a hand-built CLI with minimal automation to a
fully-gated, signed, multi-Python-tested package with auto-generated releases.
This page is the canonical reference for **what landed in each phase, why, and
where to look for it**.

If you're new to the codebase, read this top-to-bottom — every other page in
the *Contributing* section assumes the tooling described here is in place.

## Goals

The overarching goals across every phase:

1. **Make CI the source of truth.** A green pipeline should mean the change is
   safe to ship; a red pipeline should mean it is not. No "ignore that one CI
   failure, it's flaky" exceptions.
2. **Automate everything that ships to a user.** Versioning, changelog, PyPI
   upload, EXE build, SBOM generation, signature attestation — all of it
   triggered from `git merge`, never `git tag` or local `twine upload`.
3. **Keep the contributor loop fast.** One-liner commands (`just lint`,
   `just test`, `just install`) and shared editor config so a new contributor
   gets a green local check in under five minutes.
4. **Make the supply chain auditable.** Every release ships a CycloneDX SBOM
   plus a license report, both signed with Sigstore so anyone can verify
   provenance offline.

## Phase summary

| # | Phase | Focus | Commit |
| --- | --- | --- | --- |
| A | Foundation tooling | `pyproject.toml`, `.editorconfig`, `.gitattributes`, Dependabot, issue/PR templates, drop Py3.8/3.9 | `3eacf06` |
| B | Quality-gate CI | Pre-commit, ruff lint+format, mypy, pytest matrix (3.10–3.13 × Linux/Win), security scans, build-wheel, aggregate gate | `ce29ab3` |
| C | Test coverage uplift | Encoding round-trip, header processing, helper coverage, COM mocks | `88564a4` |
| D | Reliability primitives | Typed exceptions, bounded-retry helper, structured logging, more tests, coverage floor → 25% | `5815429` |
| E | COM/IO hardening | COM error handling, threading correctness, BOM-aware decoding, CLI input validation | `61981eb` |
| F | Release automation | release-please, PyPI Trusted Publishing, CycloneDX SBOM, license report, EXE attached to release | `db5265b` |
| G | UX polish | `.visiowings.toml` project config, `visiowings init` wizard, opt-out PyPI update check | `7bc6f05` |
| G+ | Docs site | MkDocs Material, mkdocstrings, GitHub Pages auto-deploy on push to `main` | `8a0b695` |
| H | Final polish | `justfile`, `noxfile.py`, Sigstore-signed SBOM/licenses, OpenSSF Scorecard, shared `.vscode/` | `19dd67e` |
| Docs | Phase H docs | `development-environment.md`, README badges, Sigstore + Scorecard docs | `826a068` |

The remainder of this page expands each phase: what problem it solved, what
files it touched, and how to interact with the result.

## Phase A — Foundation tooling and project metadata

**Problem.** The repo had a thin `setup.py`, no formatter config, no
contributor guidelines, supported EOL Pythons (3.8 and 3.9), and no
machine-readable hint about which line endings VBA files needed.

**What landed.**

- `pyproject.toml` is now the single source of truth: project metadata,
  dependencies, optional `[dev]`/`[docs]` extras, ruff/ruff-format/mypy/
  pytest/coverage config, entry points. `setup.py` is a 4-line shim kept only
  because some legacy build tools still look for it.
- `.editorconfig` enforces LF + 4-space Python and CRLF for VBA files
  (`*.bas`, `*.cls`, `*.frm`) — VBA fails to re-import if line endings drift.
- `.gitattributes` mirrors the same line-ending policy at the Git layer so
  Windows checkouts don't silently rewrite source files.
- Dependabot (`pip` + `github-actions`, weekly cadence, conventional-commit
  prefixes) so dependency PRs auto-flow.
- Issue + PR templates, `CONTRIBUTING.md`, `SECURITY.md`, `CODE_OF_CONDUCT.md`.
- Python 3.8 and 3.9 dropped (both EOL); 3.10–3.13 declared.
- `pywin32` made a Windows-only conditional dependency
  (`pywin32>=305; sys_platform == 'win32'`) so the package installs on Linux
  for tests.
- `py.typed` marker added — visiowings ships type information per PEP 561.

**Where to look.** `pyproject.toml`, `.editorconfig`, `.gitattributes`,
`.github/dependabot.yml`, `.github/ISSUE_TEMPLATE/`,
`.github/PULL_REQUEST_TEMPLATE.md`.

## Phase B — Full quality-gate CI pipeline

**Problem.** The repo had a stub CI that ran a smoke test on one Python
version. There was no formatter check, no type-check, no security scan, no
build verification, and no aggregated "all checks passed" status check
suitable for branch protection.

**What landed.**

- `.github/workflows/ci.yml` runs the following on every PR and push to `main`:
  - **Pre-commit** — runs every hook over every file, identical to local.
  - **Lint (ruff)** — `ruff check` + `ruff format --check`.
  - **Type-check (mypy)** with the per-module strict overrides.
  - **Test matrix** — pytest on Python 3.10/3.11/3.12/3.13 × Ubuntu/Windows
    with `pytest-cov`. Codecov upload on the canonical job
    (`ubuntu-latest` + 3.12).
  - **Security** — `pip-audit --strict --requirement requirements.txt`
    (vulnerability scan) + `bandit -r visiowings/ -ll` (static analysis).
  - **Dependency Review** — fails the PR if any new dependency carries a
    high-severity advisory.
  - **Build wheel + sdist** — `python -m build` + `twine check`, artefacts
    uploaded for inspection.
  - **CI gate** — a final job that aggregates the others. This is the single
    required-status-check for branch protection on `main`.
- `.github/workflows/codeql.yml` runs CodeQL on push + weekly.
- `.pre-commit-config.yaml` runs ruff, ruff-format, mypy, codespell, gitleaks
  (secret scanning), validate-pyproject, conventional-pre-commit, and
  actionlint.

**Where to look.** `.github/workflows/ci.yml`,
`.github/workflows/codeql.yml`, `.pre-commit-config.yaml`.

## Phase C — Test coverage uplift

**Problem.** The codebase had a handful of integration tests that needed a
real Visio install. There was effectively nothing testable on Linux.

**What landed.**

- `tests/_visio_mocks.py` — a typed mock harness for the pywin32 surface area
  the codebase actually uses. Lets us run almost all logic on Linux.
- New test modules covering encoding round-trips
  (`tests/test_encoding_roundtrip.py`), header processing
  (`tests/test_header_processing.py`), and Rubberduck folder integration.
- The `windows_only` pytest marker so the integration tests are easy to skip
  on non-Windows runners; CI's Linux jobs run with
  `-m "not windows_only"`.

**Where to look.** `tests/_visio_mocks.py` plus every `tests/test_*.py`.

## Phase D — Typed exceptions, retries, structured logging

**Problem.** COM errors were caught as bare `Exception` and printed; transient
failures could blow up a long-running watcher session; logging was a mix of
`print` and ad-hoc `logging.basicConfig` calls.

**What landed.**

- `visiowings/exceptions.py` — typed exception hierarchy
  (`VisiowingsError` → `COMConnectionError`, `EncodingError`,
  `VBAImportError`, `VBAExportError`).
- `visiowings/_retry.py` — bounded-retry helper with exponential backoff for
  COM operations. Deliberately no `tenacity` dependency.
- `visiowings/_logging.py` — single `setup_logging(debug: bool)` entry point
  that configures a structured root logger; respects `VISIOWINGS_LOG_LEVEL`.
- New test modules for the three primitives (`test_retry.py`,
  `test_logging_setup.py`, plus the existing exception tests).
- Coverage floor raised from 15% → 25% in `pyproject.toml`.

**Where to look.** `visiowings/exceptions.py`, `visiowings/_retry.py`,
`visiowings/_logging.py`, plus the matching tests.

## Phase E — COM, threading, encoding hardening

**Problem.** Three classes of real bugs from production use: COM channels
broke silently after Visio crashes, watchdog's threading model conflicted
with COM apartments, and BOM-prefixed files could be silently mis-decoded.

**What landed.**

- `vba_import.VisioVBAImporter._ensure_connection` — detects a dead COM
  channel via a sentinel attribute access, retries up to
  `_MAX_RECONNECT_ATTEMPTS` (3) before giving up with a typed
  `COMConnectionError`.
- `pythoncom.CoInitialize`/`CoUninitialize` calls in every watchdog handler
  thread (the import path was previously implicit, breaking on some Visio
  versions).
- `encoding.get_encoding_from_document` reads the document's LCID and maps
  it through the `LCID_TO_CODEPAGE` table; `resolve_encoding` is the new
  user-facing entry point with a clear precedence (CLI flag → document →
  system default `cp1252`).
- BOM detection in the file watcher so cross-platform editors (which often
  prepend a BOM) don't break VBA round-trips.
- CLI input validation: file existence, extension whitelist, output path
  resolution.

**Where to look.** `visiowings/vba_import.py`, `visiowings/encoding.py`,
`visiowings/file_watcher.py`, `visiowings/cli.py`.

## Phase F — Release automation and supply-chain artefacts

**Problem.** Releases were manual: bump the version in three files, write the
changelog by hand, build locally, `twine upload`. Easy to make mistakes;
no provenance.

**What landed.**

- **release-please** (`.github/workflows/release-please.yml` +
  `release-please-config.json` + `.release-please-manifest.json`).
  Reads Conventional Commits on `main`, opens or updates a Release PR that
  bumps the version everywhere it lives (`pyproject.toml`, `__init__.py`,
  `setup.py`) and rewrites `CHANGELOG.md`. Merging that PR creates the
  `vX.Y.Z` tag and GitHub Release.
- **PyPI Trusted Publishing** — `.github/workflows/publish.yml` uses an OIDC
  token to upload to PyPI. No long-lived API token. One-time PyPI setup is
  documented in [Releasing](releasing.md#one-time-setup-pypi-trusted-publishing).
- **CycloneDX SBOM** (`sbom.cdx.json`) and a **license report**
  (`licenses.json`) are generated from the wheel and attached to every
  GitHub Release.
- **Standalone Windows EXE** built via PyInstaller in
  `.github/workflows/build-exe.yml`, smoke-tested with `--version`, and
  attached to the Release alongside the wheel and sdist.

**Where to look.** `.github/workflows/publish.yml`,
`.github/workflows/release-please.yml`, `.github/workflows/build-exe.yml`,
`release-please-config.json`.

## Phase G — Project config, init wizard, update check

**Problem.** Repeat users were typing the same `--file ./mydrawing.vsdm
--bidirectional --output ./vba` flags every session. There was no way to know
a release shipped without checking PyPI by hand.

**What landed.**

- `visiowings/config.py` reads `.visiowings.toml` from the project root and
  layers it under CLI flags. Documented in
  [Configuration](../getting-started/configuration.md).
- `visiowings init` is a small interactive wizard that scaffolds a
  `.visiowings.toml` for the user. Detects open Visio documents on Windows
  and offers them as defaults.
- `visiowings/_update_check.py` does an opt-out, once-per-day check against
  PyPI's JSON API for a newer version. Cached in
  `~/.cache/visiowings/update_check.json`. Disable via
  `VISIOWINGS_NO_UPDATE_CHECK=1` or `update_check = false` in
  `.visiowings.toml`.
- Coverage floor raised → 30%.

**Where to look.** `visiowings/config.py`, `visiowings/_update_check.py`,
`visiowings/cli.py` (the `init` subcommand).

## Phase G+ — Documentation site

**Problem.** The README had grown into a 270-line wall of text. There was no
discoverable place for a "how do I configure this?" answer.

**What landed.**

- `mkdocs.yml` configures MkDocs Material with mkdocstrings (Python handler)
  for auto-generated API docs.
- `docs/` tree split into Getting Started, Architecture, API Reference,
  Contributing, Changelog.
- `.github/workflows/docs.yml` builds with `--strict` on every push and
  deploys to GitHub Pages on push to `main`. Live at
  <https://twobeass.github.io/visiowings/>.

**Where to look.** `mkdocs.yml`, `docs/`, `.github/workflows/docs.yml`.

## Phase H — Final polish: dev experience and supply-chain

**Problem.** Three small remaining gaps: contributors had to memorise long
shell incantations, the matrix tested four Pythons but you couldn't easily
reproduce that locally, and the SBOM/license artefacts shipped unsigned.

**What landed.**

- **`justfile`** — one-liner recipes for every common workflow
  (`just install`, `just test`, `just lint`, `just fmt`, `just security`,
  `just docs-serve`, …). Full table in
  [Development environment](development-environment.md#task-runner-just).
- **`noxfile.py`** — cross-Python (3.10/3.11/3.12/3.13) sessions for tests +
  lint + type-check + docs + security. Reuses virtualenvs.
- **Sigstore signing** — `publish.yml` now signs `sbom.cdx.json` and
  `licenses.json` keyless via OIDC. The `.sigstore` bundles are attached to
  the Release. Verification recipe in
  [Releasing → Verifying a release](releasing.md#verifying-a-release).
- **OpenSSF Scorecard** — `.github/workflows/scorecard.yml` runs weekly and
  on push to `main`; SARIF results land in **Security → Code scanning** for
  dashboard tracking. Public viewer link in the README badge row.
- **Shared `.vscode/`** — `settings.json` (Ruff format-on-save, mypy daemon,
  pytest discovery, sane file defaults) + `extensions.json` (recommended
  extensions). Per-user state is still ignored via
  `.gitignore` — only the two committed files are tracked.

**Where to look.** `justfile`, `noxfile.py`,
`.github/workflows/publish.yml` (Sigstore step),
`.github/workflows/scorecard.yml`, `.vscode/`.

## What didn't ship (and why)

A few things were considered and intentionally deferred:

- **Strict ruleset for `BLE`/`TRY`/`SIM`/`PL` ruff families.** Phase A
  enabled them; reality was that the legacy COM bridge in
  `vba_import.py`/`vba_export.py`/`visio_connection.py` was built around
  `except Exception` and lazy `import pywin32` patterns, neither of which
  passes those rules. Re-enabling is gated on refactoring the COM layer
  behind a typed adapter — that's bigger than a CI cleanup.
- **`disallow_untyped_defs = true` globally.** Currently scoped to the
  primitives added in Phase D (`encoding`, `_retry`, `exceptions`). The
  legacy modules use COM objects which are inherently `Any`-typed; widening
  to `Any` annotations would satisfy mypy but communicate nothing.
- **PyPI publishing of the standalone EXE.** EXEs aren't a PyPI artefact
  type; they're attached to GitHub Releases instead. The PyPI install path
  remains `pipx install visiowings`.

## Repo-side prerequisites

Two settings live in the GitHub UI (not in this repo) and are worth
calling out — without them, parts of the pipeline degrade:

- **Dependency graph** — *Settings → Code security & analysis →
  Dependency graph: Enabled*. Required by the **Dependency Review** CI
  job; without it, the action fails fast and the job is treated as
  advisory (`continue-on-error: true`).
- **PyPI Trusted Publisher** — see
  [Releasing → One-time setup](releasing.md#one-time-setup-pypi-trusted-publishing).
  Without it, the first release upload to PyPI fails; subsequent
  releases are fine once the pending publisher is promoted.

## Where the docs live now

| Topic | File |
| --- | --- |
| User-facing install + usage | `docs/getting-started/`, `README.md` |
| `.visiowings.toml` | `docs/getting-started/configuration.md` |
| Codepage matrix | `docs/getting-started/codepages.md` |
| API reference (auto-generated) | `docs/components/cli.md`, `docs/components/document-manager.md` |
| Contributor onboarding | `CONTRIBUTING.md`, `docs/contributing/development-environment.md` |
| Releasing + supply-chain verification | `docs/contributing/releasing.md` |
| Manual testing on real Visio (tutorial) | `docs/contributing/manual-testing.md` |
| Release UAT checklist (sign-off) | `docs/contributing/uat.md` |
| **This page** | `docs/contributing/optimization-plan.md` |
| Changelog | `CHANGELOG.md`, `docs/changelog.md` |
