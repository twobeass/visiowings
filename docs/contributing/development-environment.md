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
