"""Cross-Python test orchestration.

Use ``nox -s tests`` to run the test suite against every supported
Python interpreter installed on this machine, or ``nox -l`` to list
sessions.
"""

from __future__ import annotations

import nox

nox.options.reuse_existing_virtualenvs = True
nox.options.sessions = ["lint", "type_check", "tests"]

PY_VERSIONS = ("3.10", "3.11", "3.12", "3.13")


@nox.session(python=PY_VERSIONS)
def tests(session: nox.Session) -> None:
    """Run the test suite with coverage."""

    session.install("-e", ".[dev]")
    session.run("pytest", *session.posargs)


@nox.session(python="3.12")
def lint(session: nox.Session) -> None:
    """Run ruff check + ruff format --check."""

    session.install("ruff>=0.6.0")
    session.run("ruff", "check", ".")
    session.run("ruff", "format", "--check", ".")


@nox.session(python="3.12")
def type_check(session: nox.Session) -> None:
    """Run mypy with the project's strict overrides."""

    session.install("-e", ".[dev]")
    session.run("mypy", "visiowings/")


@nox.session(python="3.12")
def docs(session: nox.Session) -> None:
    """Build the mkdocs site with --strict."""

    session.install("-e", ".[docs]")
    session.run("mkdocs", "build", "--strict")


@nox.session(python="3.12", venv_backend="none")
def uat(session: nox.Session) -> None:
    """Run the in-tree UAT suite (Windows + Visio required).

    Uses the host interpreter so the COM bindings (pywin32) resolve against
    the system's registered Office installation. Auto-skips its real tests
    when Office apps are not detected.
    """

    session.run(
        "pytest",
        "tests/uat",
        "--no-cov",
        "-ra",
        *session.posargs,
        external=True,
    )


@nox.session(python="3.12")
def security(session: nox.Session) -> None:
    """Run pip-audit + bandit scans."""

    session.install("-e", ".", "pip-audit", "bandit[toml]")
    session.run("pip-audit", "--strict")
    session.run("bandit", "-r", "visiowings/", "-ll")
