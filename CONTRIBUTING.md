# Contributing to visiowings

Thanks for your interest in improving visiowings! This document covers the
development workflow, expectations for pull requests, and how to run the test
suite.

> 📖 **New here?** The
> [optimization plan (Phases A–H)](https://twobeass.github.io/visiowings/contributing/optimization-plan/)
> walks through every piece of CI, tooling, and release automation the project
> has — what it does, why it exists, and where to find it.

## Quick start

```bash
git clone https://github.com/twobeass/visiowings.git
cd visiowings
python -m venv .venv
# Windows:
.venv\Scripts\activate
# Linux/macOS:
source .venv/bin/activate

pip install -e ".[dev]"
pre-commit install
pre-commit install --hook-type commit-msg
```

If you have [`just`](https://just.systems) installed, the same setup is one
command:

```bash
just install
```

`visiowings` itself only runs on Windows (it talks to Visio via COM), but most
unit tests run on Linux and macOS thanks to mocked `pywin32` interfaces.

## Editor setup

The repo ships shared VS Code workspace settings in `.vscode/`:

- `settings.json` — Ruff format-on-save, import sorting, mypy daemon,
  pytest discovery, sane `files.*` defaults.
- `extensions.json` — recommended extensions (Ruff, Python, Pylance,
  mypy, EditorConfig, GitHub Actions, …). VS Code prompts you to install
  them on first open.

Per-user state (`.vscode/launch.json`, history, etc.) is intentionally
ignored via `.gitignore` — only the two files above are committed.

## Branching

Open a feature branch from `main`:

```bash
git switch -c feat/short-description
```

## Conventional Commits

Commit messages and PR titles must follow
[Conventional Commits](https://www.conventionalcommits.org/):

```
feat(cli): add visiowings init wizard
fix(file_watcher): debounce duplicate save events
docs: clarify codepage matrix
chore(deps): bump watchdog to 4.0
```

The release pipeline (`release-please`) reads these commits to generate
`CHANGELOG.md` and pick the next version. Use `feat!:` or
`BREAKING CHANGE:` in the body for breaking changes.

## Running checks locally

The raw commands:

```bash
# Format + lint + import-sort
ruff check --fix .
ruff format .

# Type-check
mypy visiowings/

# Tests + coverage gate
pytest

# Everything pre-commit checks (run once before pushing)
pre-commit run --all-files
```

…or, with `just`, the recipes that wrap them:

```bash
just lint      # ruff check + ruff format --check + mypy
just fmt       # ruff check --fix + ruff format
just test      # full pytest run (mirrors CI)
just test-fast # skips windows_only + slow
just security  # pip-audit --strict + bandit
just pc        # pre-commit run --all-files
just docs-serve  # live-reload docs preview
```

Run `just` (no args) to list every recipe. CI runs the same commands —
if `just lint` and `just test` are green locally, CI almost certainly is too.

### Cross-Python testing with nox

CI tests against Python 3.10–3.13. To reproduce the same matrix locally
(any interpreter you have installed will be used; missing ones are
skipped):

```bash
pip install nox     # one-time
nox                 # default sessions: lint + type_check + tests
nox -s tests        # tests on every available 3.10–3.13
nox -s tests -- -k encoding   # forward args to pytest
nox -l              # list every session
```

## Testing on Windows with real Visio

Tests marked `@pytest.mark.windows_only` need a real Visio instance:

```bash
pytest -m windows_only
```

Skip them with `pytest -m "not windows_only"` (the default in CI for the
non-Windows runners).

## Adding tests

- Unit tests live in `tests/`.
- COM interactions are mocked via the helpers in `tests/_visio_mocks.py`.
- Encoding-sensitive code paths must come with a round-trip test in
  `tests/test_encoding_roundtrip.py`.

## Releasing

You don't need to do anything to cut a release. After your PR is merged to
`main`, `release-please` will open (or update) a Release PR that bumps the
version and updates `CHANGELOG.md`. Merging that Release PR triggers PyPI
publication and a GitHub Release with the standalone `.exe`.

## Reporting bugs and proposing features

Use the issue templates in `.github/ISSUE_TEMPLATE/`. For security issues,
please file a private advisory: see [SECURITY.md](SECURITY.md).
