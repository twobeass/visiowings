# `just` recipes for visiowings — install with https://just.systems
# Usage: `just <recipe>` (or `just` to list).

# List all recipes
default:
    @just --list

# Create venv and install all dev + docs dependencies, then enable hooks
install:
    python -m pip install --upgrade pip
    pip install -e ".[dev,docs]"
    pre-commit install
    pre-commit install --hook-type commit-msg

# Run the full test suite (mirrors CI)
test:
    pytest

# Run tests with the same JUnit + coverage XML CI produces
test-ci:
    pytest --cov-report=xml --junitxml=junit.xml

# Run only fast tests (skip windows_only + slow markers)
test-fast:
    pytest -m "not windows_only and not slow" -x

# Lint + format + type-check (read-only)
lint:
    ruff check .
    ruff format --check .
    mypy visiowings/

# Auto-format + auto-fix lint
fmt:
    ruff check --fix .
    ruff format .

# Security scans (pip-audit + bandit)
security:
    pip-audit --strict
    bandit -r visiowings/ -ll

# Run all pre-commit hooks against every file
pc:
    pre-commit run --all-files --show-diff-on-failure

# Build the docs site locally (with watch reload)
docs-serve:
    mkdocs serve

# Build a static docs site (CI uses --strict)
docs-build:
    mkdocs build --strict

# Build wheel + sdist
build-wheel:
    python -m build
    twine check --strict dist/*

# Build the standalone Windows EXE (Windows only)
build-exe:
    pyinstaller visiowings.spec --clean --noconfirm

# Bump the coverage gate floor; first arg = new percentage
bump-cov FLOOR:
    sed -i 's/^fail_under = .*/fail_under = {{FLOOR}}/' pyproject.toml

# Wipe build artefacts
clean:
    rm -rf build dist site coverage.xml junit.xml htmlcov .pytest_cache .mypy_cache .ruff_cache .coverage .coverage.* sbom.cdx.json licenses.json
    find . -type d -name __pycache__ -exec rm -rf {} +

# Print version + CLI help
info:
    visiowings --version
    visiowings --help
