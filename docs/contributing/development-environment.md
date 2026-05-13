# Development environment

This page is the one-stop reference for the tooling that surrounds the
codebase: task runners, cross-Python testing, pre-commit hooks, and
shared editor configuration. It's optional reading if you already use
`pip install -e ".[dev]"` and `pytest`, but recommended if you plan to
contribute regularly.

## Task runner: `just`

[`just`](https://just.systems) is a small command runner. Recipes live
in the [`justfile`](https://github.com/twobeass/visiowings/blob/main/justfile)
at the repo root. Type `just` (no args) for the full list.

| Recipe | What it does |
| --- | --- |
| `just install` | Create venv-friendly install (`pip install -e ".[dev,docs]"`) and enable pre-commit hooks. |
| `just test` | Full pytest suite (mirrors CI). |
| `just test-ci` | Same, plus JUnit + coverage XML artefacts. |
| `just test-fast` | Skips `windows_only` and `slow` markers, fails fast. |
| `just test-uat` | Run the in-tree UAT suite (`tests/uat/`) — Windows + Visio required, auto-skips otherwise. |
| `just uat-fixtures` | Generate `fixtures/sample.vsdm` for the UAT suite. |
| `just uat-trust-center` | One-shot Trust Center bootstrap (HKCU `AccessVBOM`=1). |
| `just lint` | `ruff check` + `ruff format --check` + `mypy` (read-only). |
| `just fmt` | `ruff check --fix` + `ruff format` (auto-fix). |
| `just security` | `pip-audit --strict` + `bandit -ll`. |
| `just pc` | Run every pre-commit hook against every file. |
| `just docs-serve` | Live-reload MkDocs preview at <http://127.0.0.1:8000>. |
| `just docs-build` | Build the static site with `--strict`. |
| `just build-wheel` | `python -m build` + `twine check --strict`. |
| `just build-exe` | PyInstaller standalone EXE (Windows only). |
| `just bump-cov FLOOR` | Raise the coverage gate floor in `pyproject.toml`. |
| `just clean` | Wipe build artefacts and caches. |
| `just info` | Print `visiowings --version` and `--help`. |

If you don't have `just`, every recipe is a thin wrapper around plain
shell commands — copy the body from the `justfile` and run it directly.

## Cross-Python testing: `nox`

CI tests against Python 3.10–3.13. The
[`noxfile.py`](https://github.com/twobeass/visiowings/blob/main/noxfile.py)
reproduces the same matrix locally, using whichever interpreters are
installed on your machine (others are skipped).

```bash
pip install nox     # one-time
nox                 # default: lint + type_check + tests
nox -s tests        # tests on every available 3.10–3.13
nox -s tests-3.12   # only one interpreter
nox -s tests -- -k encoding   # forward args to pytest
nox -s docs         # mkdocs build --strict
nox -s security     # pip-audit + bandit
nox -l              # list every session
```

`nox` reuses virtualenvs across runs (`reuse_existing_virtualenvs = True`),
so the second invocation is fast.

!!! tip "Installing multiple Pythons"
    Use [`pyenv`](https://github.com/pyenv/pyenv) on Linux/macOS or
    [`uv python install`](https://docs.astral.sh/uv/) on any platform to
    grab missing interpreters. Nox automatically finds them on `PATH`.

## UAT automation: `tests/uat/`

The repo ships a **pytest-driven UAT suite** that exercises the real
`visiowings` CLI against a live Visio instance. Each test maps 1:1 to a
section of [Human UAT](uat.md), so the markdown stays the source of truth
and the automation backs it up.

Layout:

| Path | Purpose |
| --- | --- |
| `tests/uat/test_visiowings_uat.py` | One test per UAT section (§A–§L). |
| `tests/uat/_helpers.py` | `run_branch`, `visiowings_cli`, `WatcherHandle`, `start_watcher`. |
| `tests/uat/com_helpers/` | Visio / VBE / process zombie cleanup helpers. |
| `tests/uat/setup/` | One-shot `bootstrap` (`fixture_factory`, `trust_center`, `office_detect`). |
| `tests/uat/conftest.py` | Scoped fixtures + auto-skip when Office / Visio / network is unavailable. |
| `tests/uat/markers.py` | `section`, `requires_office`, `manual_signoff`, `not_yet_implemented`, … |

Bootstrap & run:

```powershell
pip install -e ".[uat]"                  # psutil, pywin32, pytest-timeout, …
python -m tests.uat.setup.trust_center   # one-shot HKCU AccessVBOM=1
python -m tests.uat.setup.fixture_factory  # creates fixtures/sample.vsdm
pytest tests/uat --no-cov                # or `just test-uat`
```

A handful of sections (§C2/§D1-§D5/§E1/§E4) need a **user-opened** Visio
with `fixtures/sample.vsdm` because cross-process `Dispatch("Visio.Application")`
spawns a fresh instance on this Office build. The tests detect that
prerequisite and skip with explicit instructions rather than failing.

> The UAT suite was previously hosted in a sibling `vbatest` repo that
> orchestrated both `visiowings` and `VBAlidator`. The visiowings-specific
> half has been folded back into this repo so the suite no longer requires
> a parallel checkout. The external `vbatest` orchestrator is now optional
> and only useful for cross-repo (visiowings × VBAlidator) testing.

## Pre-commit hooks

The `.pre-commit-config.yaml` runs Ruff, mypy, codespell, conventional
commit linting, and a handful of safety checks. Install it once:

```bash
pre-commit install
pre-commit install --hook-type commit-msg   # for conventional-commit linting
```

To run every hook against every file (matches CI):

```bash
pre-commit run --all-files
# or
just pc
```

## Editor: shared `.vscode/` config

The repo ships two opt-in workspace files:

- **`.vscode/settings.json`** — Ruff format-on-save, organize-imports,
  mypy daemon, pytest discovery, sane `files.*` defaults, and `.bas /
  .cls / .frm` association as VB.
- **`.vscode/extensions.json`** — recommended extensions; VS Code
  prompts you to install them on first open.

Per-user files (`launch.json`, history, etc.) are intentionally ignored
in `.gitignore` — only the two files above are checked in. If you use
JetBrains, Vim, or another editor, the project is also fully driven by
[`.editorconfig`](https://github.com/twobeass/visiowings/blob/main/.editorconfig)
and `pyproject.toml`, so formatting parity is guaranteed across editors.

## Supply-chain tooling

These are mostly invisible during day-to-day development, but worth
knowing about:

- **CycloneDX SBOM** — generated on every release and attached to the
  GitHub Release.
- **Sigstore signatures** — the SBOM and license report are signed
  keyless via OIDC; verification instructions live in
  [Releasing → Verifying a release](releasing.md#verifying-a-release).
- **OpenSSF Scorecard** — runs weekly and on every push to `main`;
  results land in the repo's **Security → Code scanning** tab.
- **PyPI Trusted Publishing** — no long-lived API tokens; PyPI accepts
  uploads only from `publish.yml` running on a tagged release.

See [Releasing](releasing.md) for the full release pipeline.
