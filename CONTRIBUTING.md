# Contributing to visiowings

Thanks for your interest in improving visiowings! This document covers the
development workflow, expectations for pull requests, and how to run the test
suite.

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

`visiowings` itself only runs on Windows (it talks to Visio via COM), but most
unit tests run on Linux and macOS thanks to mocked `pywin32` interfaces.

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

CI runs the same commands. If `pre-commit run --all-files` is green locally,
CI is almost certainly going to be green too.

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
