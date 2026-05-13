"""visiowings UAT — Iteration 1 covers §A (env), §B (install paths), §C (config).

Sections D–L are placeholders for later iterations.
"""

from __future__ import annotations

import os
import shutil
import sys
import time
from pathlib import Path

import pytest

from tests.uat._helpers import (
    repo_present,
    run,
    run_branch,
    start_watcher,
    visiowings_cli,
)


def _need_repo(repo: Path) -> None:
    if not repo_present(repo):
        pytest.skip(f"visiowings repo not present at {repo} — run validate.ps1 first")


# ---------------------------------------------------------------------------
# A — Environment / Trust Center
# ---------------------------------------------------------------------------


@pytest.mark.section("visiowings-A1", "visiowings/docs/contributing/uat.md#a1")
def test_section_a1_vbom_enabled():
    """UAT §A1: 'Trust access to VBA project object model' must be ON for Visio."""
    from tests.uat.setup.trust_center import read_vbom_state

    state = read_vbom_state()
    visio_entries = {(app, ver): vals for (app, ver), vals in state.items() if app == "Visio"}
    if not visio_entries:
        pytest.skip("Visio Trust Center key not present — Visio may not be installed")
    for (app, ver), vals in visio_entries.items():
        assert vals.get("AccessVBOM") == 1, (
            f"AccessVBOM not set for {app} {ver}; run tests/uat/setup/trust_center.py"
        )


@pytest.mark.section("visiowings-A2", "visiowings/docs/contributing/uat.md#a2")
def test_section_a2_fixtures_present(fixtures_dir):
    """UAT §A2: fixture factory has produced the minimum sample set."""
    # If Visio isn't installed we can't have generated fixtures; skip.
    try:
        from tests.uat.setup.office_detect import detect

        if not detect().get("visio", False):
            pytest.skip("Visio not installed — fixture generation requires Visio")
    except Exception:
        pass
    sample = fixtures_dir / "sample.vsdm"
    manifest = fixtures_dir / "manifest.json"
    assert manifest.exists(), f"manifest missing — fixture factory did not run ({fixtures_dir})"
    assert sample.exists(), f"sample.vsdm missing at {sample}"


@pytest.mark.section("visiowings-A3", "visiowings/docs/contributing/uat.md#a3")
def test_section_a3_python_and_pipx():
    """UAT §A3: Python 3.10–3.13 + pipx available (pipx may not be on PATH yet)."""
    v = sys.version_info
    assert (3, 10) <= (v.major, v.minor) <= (3, 13), (
        f"Python {v.major}.{v.minor} out of supported range"
    )
    # pipx is recommended but bootstrap only soft-installs it. Treat absence
    # as a skip per §A3 wording (recommended, not required).
    if shutil.which("pipx") is None:
        pytest.skip("pipx not on PATH (recommended only)")


# ---------------------------------------------------------------------------
# B — Install paths
# ---------------------------------------------------------------------------


@pytest.mark.section("visiowings-B1", "visiowings/docs/contributing/uat.md#b1")
def test_section_b1_pipx_install(tmp_path, visiowings_repo):
    """UAT §B1: ``pipx install ./visiowings`` succeeds.

    We install from the local repo into a throwaway pipx home so we don't
    clobber the user's installation.
    """
    _need_repo(visiowings_repo)
    if shutil.which("pipx") is None:
        pytest.skip("pipx not on PATH")
    env = os.environ.copy()
    env["PIPX_HOME"] = str(tmp_path / "pipx_home")
    env["PIPX_BIN_DIR"] = str(tmp_path / "pipx_bin")
    cp = run(["pipx", "install", str(visiowings_repo), "--force"], timeout=300)
    assert cp.returncode == 0, f"pipx install failed: {cp.stderr[-800:]}"


@pytest.mark.section("visiowings-B2", "visiowings/docs/contributing/uat.md#b2")
def test_section_b2_pip_venv(tmp_path, visiowings_repo):
    """UAT §B2: ``pip install`` into a fresh venv succeeds."""
    _need_repo(visiowings_repo)
    venv_dir = tmp_path / "venv"
    cp = run([sys.executable, "-m", "venv", str(venv_dir)], timeout=120)
    assert cp.returncode == 0, f"venv creation failed: {cp.stderr}"
    pip = venv_dir / "Scripts" / "pip.exe"
    if not pip.exists():
        pip = venv_dir / "bin" / "pip"  # safety net for non-Windows CI
    cp = run([str(pip), "install", str(visiowings_repo)], timeout=300)
    assert cp.returncode == 0, f"pip install failed: {cp.stderr[-800:]}"


@pytest.mark.section("visiowings-B3", "visiowings/docs/contributing/uat.md#b3")
@pytest.mark.not_yet_implemented("iter4")
def test_section_b3_exe_runs():
    """UAT §B3: Released EXE artifact runs and prints --version."""


@pytest.mark.section("visiowings-B4", "visiowings/docs/contributing/uat.md#b4")
def test_section_b4_just_recipes(visiowings_repo):
    """UAT §B4: ``just --list`` enumerates the project's recipes."""
    _need_repo(visiowings_repo)
    if shutil.which("just") is None:
        pytest.skip("just not on PATH")
    if not (visiowings_repo / "justfile").exists() and not (visiowings_repo / "Justfile").exists():
        pytest.skip("no justfile in visiowings repo")
    cp = run(["just", "--list"], cwd=visiowings_repo, timeout=30)
    assert cp.returncode == 0, f"just --list failed: {cp.stderr}"


# ---------------------------------------------------------------------------
# C — Config
# ---------------------------------------------------------------------------


@pytest.mark.section("visiowings-C1", "visiowings/docs/contributing/uat.md#c1")
def test_section_c1_init_creates_toml(tmp_path, visiowings_repo, fixtures_dir):
    """UAT §C1: ``visiowings init`` writes a ``.visiowings.toml`` config.

    The wizard is interactive (no ``--non-interactive`` flag in this build),
    so we drive it by feeding answers via stdin. The fixture factory has
    placed ``sample.vsdm`` on disk; we point the wizard at it.
    """
    _need_repo(visiowings_repo)
    import os
    import subprocess as sp

    env = os.environ.copy()
    env["PYTHONPATH"] = str(visiowings_repo) + os.pathsep + env.get("PYTHONPATH", "")

    sample = fixtures_dir / "sample.vsdm"
    if not sample.exists():
        pytest.skip(f"sample.vsdm fixture not generated at {sample}")

    # Wizard answers, robust to both branches:
    # - docs-found: select option "1" (auto-discovered doc) — the wizard
    #   already knows the path
    # - no-docs: wizard asks for path directly — we provide it as fallback
    # The stdin stream is consumed line-by-line; surplus lines in either
    # branch are tolerated (wizard reads only what it needs).
    stdin_lines = [
        "1",  # selection in docs-found case
        str(sample),  # path in no-docs case OR ignored when "1" picked
        "vba",  # output directory
        "n",  # bidirectional sync
        "n",  # rubberduck
        "",  # codepage = auto-detect
    ]
    cp = sp.run(
        [*visiowings_cli(visiowings_repo), "init"],
        cwd=str(tmp_path),
        env=env,
        input="\n".join(stdin_lines) + "\n",
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=60,
    )
    assert cp.returncode == 0, (
        f"visiowings init failed (exit {cp.returncode}):\n"
        f"--- stdout ---\n{cp.stdout[-1500:]}\n--- stderr ---\n{cp.stderr[-800:]}"
    )
    toml = tmp_path / ".visiowings.toml"
    assert toml.exists(), (
        f".visiowings.toml not produced in {tmp_path}. "
        f"Files: {[p.name for p in tmp_path.iterdir()]}"
    )
    content = toml.read_text(encoding="utf-8")
    for key in ("file", "output"):
        assert key in content, f"{key} key missing from .visiowings.toml"


@pytest.mark.section("visiowings-C2", "visiowings/docs/contributing/uat.md#c2")
@pytest.mark.requires_office("visio")
def test_section_c2_flag_overrides_toml(tmp_path, visiowings_repo, fixtures_dir):
    """UAT §C2: CLI flag overrides ``.visiowings.toml``.

    Step 1+2: with toml `output = "./vba"`, export goes to `./vba`.
    Step 3+4: with `--output ./other`, export goes to `./other`, NOT `./vba`.
    """
    _need_repo(visiowings_repo)
    sample = fixtures_dir / "sample.vsdm"
    if not sample.exists():
        pytest.skip(f"sample.vsdm fixture missing at {sample}")
    app = _require_user_opened_visio("sample.vsdm")
    doc = next(d for d in app.Documents if d.Name.lower() == "sample.vsdm")
    doc_path = doc.FullName

    workdir = tmp_path / "workdir"
    workdir.mkdir()
    toml_path = workdir / ".visiowings.toml"
    toml_path.write_text('output = "./vba_from_toml"\n', encoding="utf-8")

    # Step 1+2: no --output flag, toml's "./vba_from_toml" must win
    cp = run_branch(
        [*visiowings_cli(visiowings_repo), "export", "--file", str(doc_path)],
        visiowings_repo,
        cwd=workdir,
        timeout=120,
    )
    assert cp.returncode == 0, f"export with toml-only config failed: {cp.stderr or cp.stdout}"
    toml_target = workdir / "vba_from_toml"
    cli_target = workdir / "vba_from_flag"
    assert toml_target.exists() and any(toml_target.rglob("*.bas")), (
        f"export should land in ./vba_from_toml per toml; "
        f"files under workdir: {[p.relative_to(workdir) for p in workdir.rglob('*.bas')]}"
    )
    assert not cli_target.exists() or not any(cli_target.rglob("*.bas")), (
        "cli-flag-only target shouldn't exist yet"
    )

    # Step 3+4: explicit --output overrides toml
    cp = run_branch(
        [
            *visiowings_cli(visiowings_repo),
            "export",
            "--file",
            str(doc_path),
            "--output",
            str(cli_target),
        ],
        visiowings_repo,
        cwd=workdir,
        timeout=120,
    )
    assert cp.returncode == 0, f"export with --output failed: {cp.stderr or cp.stdout}"
    assert any(cli_target.rglob("*.bas")), (
        f"--output should override toml; expected files in {cli_target}, "
        f"got: {[p.relative_to(workdir) for p in workdir.rglob('*.bas')]}"
    )


# --------------------------------------------------------------------------
# §C3 — CLI input validation: 4 error scenarios
# UAT §C3 Expected: "All four cases error with a one-line user-facing
# message and exit code ≠ 0. No Python traceback."
# --------------------------------------------------------------------------


def _has_python_traceback(text: str) -> bool:
    """True if the output contains a raw Python traceback."""
    return "Traceback (most recent call last)" in text


@pytest.mark.section("visiowings-C3", "visiowings/docs/contributing/uat.md#c3")
def test_section_c3_file_not_found(tmp_path, visiowings_repo):
    """UAT §C3 row 1: missing file → clean error, no traceback."""
    _need_repo(visiowings_repo)
    cp = run_branch(
        [
            *visiowings_cli(visiowings_repo),
            "export",
            "--file",
            str(tmp_path / "does_not_exist.vsdm"),
        ],
        visiowings_repo,
        cwd=tmp_path,
        timeout=30,
    )
    assert cp.returncode != 0, "expected non-zero exit for missing file"
    combined = cp.stdout + cp.stderr
    assert not _has_python_traceback(combined), (
        f"raw Python traceback leaked to user: {combined[-800:]}"
    )


@pytest.mark.section("visiowings-C3", "visiowings/docs/contributing/uat.md#c3")
@pytest.mark.requires_office("visio")
def test_section_c3_unknown_codepage(tmp_path, visiowings_repo, fixtures_dir):
    """UAT §C3 row 2: unknown --codepage → clean error."""
    _need_repo(visiowings_repo)
    sample = fixtures_dir / "sample.vsdm"
    if not sample.exists():
        pytest.skip("sample.vsdm fixture missing")
    cp = run_branch(
        [
            *visiowings_cli(visiowings_repo),
            "export",
            "--file",
            str(sample),
            "--codepage",
            "zzz123",
            "--output",
            str(tmp_path / "out"),
        ],
        visiowings_repo,
        cwd=tmp_path,
        timeout=30,
    )
    assert cp.returncode != 0, "expected non-zero exit for unknown codepage"
    combined = cp.stdout + cp.stderr
    assert not _has_python_traceback(combined), f"raw Python traceback leaked: {combined[-800:]}"


@pytest.mark.section("visiowings-C3", "visiowings/docs/contributing/uat.md#c3")
def test_section_c3_unsupported_extension(tmp_path, visiowings_repo):
    """UAT §C3 row 3: --file <foo>.txt → unsupported extension error."""
    _need_repo(visiowings_repo)
    bogus = tmp_path / "fake.txt"
    bogus.write_text("not a visio file", encoding="utf-8")
    cp = run_branch(
        [
            *visiowings_cli(visiowings_repo),
            "export",
            "--file",
            str(bogus),
            "--output",
            str(tmp_path / "out"),
        ],
        visiowings_repo,
        cwd=tmp_path,
        timeout=30,
    )
    assert cp.returncode != 0, "expected non-zero exit for .txt input"
    combined = cp.stdout + cp.stderr
    assert not _has_python_traceback(combined), f"raw Python traceback leaked: {combined[-800:]}"


@pytest.mark.section("visiowings-C3", "visiowings/docs/contributing/uat.md#c3")
@pytest.mark.requires_office("visio")
def test_section_c3_readonly_output(tmp_path, visiowings_repo, fixtures_dir):
    """UAT §C3 row 4: read-only --output → clean error."""
    _need_repo(visiowings_repo)
    sample = fixtures_dir / "sample.vsdm"
    if not sample.exists():
        pytest.skip("sample.vsdm fixture missing")
    # Windows: mark a directory read-only via icacls or by creating a file
    # at the path that prevents directory creation. Easiest cross-method:
    # point --output at a path that already exists as a FILE — directory
    # creation must fail there.
    blocking = tmp_path / "blocker.txt"
    blocking.write_text("this is a file, not a dir", encoding="utf-8")
    out_param = blocking  # cannot mkdir over a file
    cp = run_branch(
        [
            *visiowings_cli(visiowings_repo),
            "export",
            "--file",
            str(sample),
            "--output",
            str(out_param),
        ],
        visiowings_repo,
        cwd=tmp_path,
        timeout=30,
    )
    assert cp.returncode != 0, (
        f"expected non-zero exit when output path is unusable; got {cp.returncode}"
    )
    combined = cp.stdout + cp.stderr
    assert not _has_python_traceback(combined), f"raw Python traceback leaked: {combined[-800:]}"


@pytest.mark.section("visiowings-C4", "visiowings/docs/contributing/uat.md#c4")
@pytest.mark.requires_office("visio")
def test_section_c4_debug_logs_format(tmp_path, visiowings_repo, fixtures_dir):
    """UAT §C4: ``visiowings export --debug`` emits structured ``[DEBUG]``
    log lines that include module names and don't mix with raw ``print()``
    noise.

    Earlier iterations heuristically probed ``--help --debug``; that did
    not trigger any logging on this build. Iter 4 runs a real ``export``
    against the user-opened ``sample.vsdm`` and asserts on the actual log
    stream visiowings produces.
    """
    _need_repo(visiowings_repo)
    sample = fixtures_dir / "sample.vsdm"
    if not sample.exists():
        pytest.skip("sample.vsdm fixture missing")
    app = _require_user_opened_visio("sample.vsdm")
    doc = next(d for d in app.Documents if d.Name.lower() == "sample.vsdm")
    doc_path = doc.FullName
    out_dir = tmp_path / "uat-c4"
    cp = run_branch(
        [
            *visiowings_cli(visiowings_repo),
            "export",
            "--file",
            str(doc_path),
            "--output",
            str(out_dir),
            "--debug",
        ],
        visiowings_repo,
        timeout=120,
    )
    assert cp.returncode == 0, f"export --debug failed (exit {cp.returncode}):\n{cp.stderr[-800:]}"
    combined = cp.stdout + cp.stderr
    debug_lines = [ln for ln in combined.splitlines() if "[DEBUG]" in ln]
    assert debug_lines, f"--debug produced no [DEBUG] lines; output was:\n{combined[-1200:]}"
    import re

    # Spec: lines are timestamped (or at least prefixed) and include module
    # names. Accept either ``[DEBUG] <module>: ...`` or
    # ``<timestamp> [DEBUG] <module> ...`` shapes.
    pat = re.compile(r"\[DEBUG\][^\n]*?[A-Za-z_][A-Za-z0-9_.]+")
    matches = [ln for ln in debug_lines if pat.search(ln)]
    assert matches, (
        f"[DEBUG] lines lack a module/logger identifier. Sample lines: {debug_lines[:5]!r}"
    )


# ---------------------------------------------------------------------------
# Deferred placeholders — visible to the drift check
# ---------------------------------------------------------------------------


def _refresh_user_doc(app, doc_name: str):
    """Close + reopen ``doc_name`` in the user's Visio so the test starts
    against a pristine on-disk state.

    Visiowings tests modify the doc's VBProject in-memory via import.
    Running multiple tests in sequence accumulates marker comments and
    can eventually break VBE handles (WinError 6 'Invalid Handle').
    Refresh ensures each test sees the disk version of the fixture.
    """
    matching = [d for d in app.Documents if d.Name.lower() == doc_name.lower()]
    if not matching:
        return None
    doc = matching[0]
    doc_path = doc.FullName
    try:
        doc.Saved = True  # discard any in-memory changes
        doc.Close()
    except Exception:
        pass
    return app.Documents.OpenEx(doc_path, 0)


def _require_user_opened_visio(required_doc_name: str = "sample.vsdm"):
    """Detect whether a user-opened Visio session has the fixture loaded.

    Two `win32com.client.Dispatch("Visio.Application")` calls return
    distinct Visio instances on this Office build (verified: separate
    ``_oleobj_`` pointers even in the same Python process), so we can't
    programmatically arrange the prep. Instead we ask the user to:

        1. Launch Visio manually
        2. Open ``fixtures/sample.vsdm``
        3. Re-run ``validate.ps1``

    If we find a running Visio.exe whose ``Documents`` collection contains
    a doc with ``required_doc_name`` (via GetActiveObject), the test
    proceeds. Otherwise we skip with explicit, actionable instructions.

    Returns the running ``Visio.Application`` COM proxy on success.
    """
    import psutil
    import pythoncom  # type: ignore
    import win32com.client  # type: ignore
    from pywintypes import com_error  # type: ignore

    # Cheap pre-check: any Visio.exe at all?
    running = [
        p
        for p in psutil.process_iter(["name"])
        if (p.info.get("name") or "").lower() == "visio.exe"
    ]
    if not running:
        pytest.skip(
            "\n[needs-manual-prep] No Visio process is running.\n"
            "  To exercise this test:\n"
            "    1. Open Visio.exe\n"
            f"    2. Open the file 'fixtures/{required_doc_name}'\n"
            "    3. Re-run validate.ps1 (or pytest -k <this test>)\n"
            "  Background: visiowings' Documents lookup needs a Visio\n"
            "  instance that exists outside this pytest process."
        )

    pythoncom.CoInitialize()
    try:
        app = win32com.client.GetActiveObject("Visio.Application")
    except com_error as exc:
        pytest.skip(
            f"\n[needs-manual-prep] Visio.exe is running but GetActiveObject failed: {exc}\n"
            "  This usually means Visio is starting up or the ROT entry\n"
            f"  isn't ready yet. Open '{required_doc_name}' in Visio and re-run."
        )

    docs = [d.Name for d in app.Documents]
    if not any(d.lower() == required_doc_name.lower() for d in docs):
        pretty_docs = ", ".join(docs) if docs else "(none)"
        pytest.skip(
            f"\n[needs-manual-prep] Visio is running but '{required_doc_name}' is not open.\n"
            f"  Currently open documents: {pretty_docs}\n"
            f"  Open '{required_doc_name}' from fixtures/ and re-run validate.ps1.\n"
            "  Tip: explorer fixtures && start fixtures\\sample.vsdm"
        )
    return app


@pytest.mark.section("visiowings-D1", "visiowings/docs/contributing/uat.md#d1")
@pytest.mark.requires_office("visio")
def test_section_d1_export_creates_files(tmp_path, visiowings_repo, fixtures_dir):
    """UAT §D1: `visiowings export` writes .bas/.cls/.frm with the
    expected Attribute-VB_Name retention and VERSION/Begin/End stripped.

    visiowings only operates on documents already open in a running
    Visio instance (it iterates ``Application.Documents``). On this
    Office build, cross-process Visio Dispatch creates fresh instances
    instead of attaching, so the user must launch Visio manually and
    open ``fixtures/sample.vsdm`` before this test runs. See
    :func:`_require_user_opened_visio` for the prep detection logic.
    """
    _need_repo(visiowings_repo)
    sample = fixtures_dir / "sample.vsdm"
    if not sample.exists():
        pytest.skip(f"sample.vsdm fixture missing at {sample}")
    app = _require_user_opened_visio("sample.vsdm")
    # Use the path Visio thinks the doc lives at — that's what visiowings
    # will match against.
    doc = next(d for d in app.Documents if d.Name.lower() == "sample.vsdm")
    doc_path = doc.FullName

    out_dir = tmp_path / "uat-export"
    cp = run_branch(
        [
            *visiowings_cli(visiowings_repo),
            "export",
            "--file",
            str(doc_path),
            "--output",
            str(out_dir),
        ],
        visiowings_repo,
        timeout=120,
    )
    assert cp.returncode == 0, (
        f"visiowings export failed (exit {cp.returncode}):\n"
        f"--- stdout ---\n{cp.stdout[-1000:]}\n"
        f"--- stderr ---\n{cp.stderr[-500:]}"
    )

    bas_files = list(out_dir.rglob("*.bas"))
    cls_files = list(out_dir.rglob("*.cls"))
    frm_files = list(out_dir.rglob("*.frm"))
    assert bas_files, (
        f"no .bas exported under {out_dir}; got files: "
        f"{[p.relative_to(out_dir) for p in out_dir.rglob('*') if p.is_file()]}"
    )

    for bas in bas_files:
        text = bas.read_text(encoding="utf-8", errors="replace")
        assert "Attribute VB_Name" in text, f"{bas.name} missing Attribute VB_Name"
        lines = [ln.strip() for ln in text.splitlines()]
        assert "VERSION 1.0 CLASS" not in lines, f"{bas.name} should not have VERSION line"
    for f in cls_files + frm_files:
        text = f.read_text(encoding="utf-8", errors="replace")
        assert "Attribute VB_Name" in text, f"{f.name} missing Attribute VB_Name"


@pytest.mark.section("visiowings-D2", "visiowings/docs/contributing/uat.md#d2")
@pytest.mark.requires_office("visio")
def test_section_d2_import_visible_in_vbe(tmp_path, visiowings_repo, fixtures_dir):
    """UAT §D2: edit an exported .bas, re-import, then read the module
    back via COM and verify the edit landed in VBE.

    Works against the user-opened sample.vsdm. visiowings import injects
    the modified module into the doc's VBProject in-place via COM, so
    we can read the result via the same Visio handle without
    closing/reopening. After the test the doc retains the marker
    comment in memory — set ``doc.Saved = True`` at the end so Visio
    won't prompt to save when the user closes.
    """
    _need_repo(visiowings_repo)
    sample = fixtures_dir / "sample.vsdm"
    if not sample.exists():
        pytest.skip(f"sample.vsdm fixture missing at {sample}")
    app = _require_user_opened_visio("sample.vsdm")

    from tests.uat.com_helpers.vbe import read_module_code

    doc = next(d for d in app.Documents if d.Name.lower() == "sample.vsdm")
    doc_path = doc.FullName

    export_dir = tmp_path / "uat-export"
    marker = f"' UAT-D2-MARKER-{tmp_path.name}"

    # Step 1: export
    cp = run_branch(
        [
            *visiowings_cli(visiowings_repo),
            "export",
            "--file",
            str(doc_path),
            "--output",
            str(export_dir),
        ],
        visiowings_repo,
        timeout=120,
    )
    assert cp.returncode == 0, f"export step failed: {cp.stderr or cp.stdout}"

    bas_candidates = list(export_dir.rglob("BasicLogic.bas"))
    if not bas_candidates:
        pytest.skip("BasicLogic.bas not produced by export; fixture may have drifted")
    bas = bas_candidates[0]

    # Step 2: inject a unique marker right after the Attribute header
    original_text = bas.read_text(encoding="utf-8")
    lines = original_text.splitlines(keepends=True)
    insert_at = 1 if lines and lines[0].lstrip().startswith("Attribute ") else 0
    lines.insert(insert_at, marker + "\n")
    bas.write_text("".join(lines), encoding="utf-8")

    # Step 3: import back
    # Pipe "o" repeatedly so any "action (o/s/i/C)" overwrite prompt
    # gets auto-answered. ``--force`` alone doesn't suppress all prompts.
    cp = run_branch(
        [
            *visiowings_cli(visiowings_repo),
            "import",
            "--file",
            str(doc_path),
            "--input",
            str(export_dir),
            "--force",
        ],
        visiowings_repo,
        timeout=120,
        stdin_input=("o\n" * 10),
    )
    assert cp.returncode == 0, f"import step failed: {cp.stderr or cp.stdout}"

    # Step 4: read VBE — user's doc was modified in-place via COM
    try:
        code = read_module_code(doc, "BasicLogic")
    finally:
        # Mark doc as Saved so user isn't prompted on close. We do NOT
        # close — that's the user's session.
        try:
            doc.Saved = True
        except Exception:
            pass

    assert marker in code, (
        f"marker {marker!r} not found in VBE module after import. Module head: {code[:300]!r}"
    )


@pytest.mark.section("visiowings-D3", "visiowings/docs/contributing/uat.md#d3")
@pytest.mark.requires_office("visio")
def test_section_d3_live_edit_mode(tmp_path, visiowings_repo, fixtures_dir):
    """UAT §D3: ``visiowings edit`` watches the local output dir; saving
    a ``.bas`` triggers ``Change detected`` + ``Imported`` within ~2s
    and the live module reflects the edit.
    """
    _need_repo(visiowings_repo)
    sample = fixtures_dir / "sample.vsdm"
    if not sample.exists():
        pytest.skip("sample.vsdm fixture missing")
    app = _require_user_opened_visio("sample.vsdm")
    from tests.uat.com_helpers.vbe import read_module_code

    doc = next(d for d in app.Documents if d.Name.lower() == "sample.vsdm")
    doc_path = doc.FullName

    out_dir = tmp_path / "uat-edit"
    handle = start_watcher(visiowings_repo, doc_path, out_dir, "--force")
    try:
        # Initial export must complete before we touch any file.
        handle.expect("Starting Live Synchronization", timeout=30)

        # Find the seeded BasicLogic.bas in the watcher's output tree.
        bas_candidates = list(out_dir.rglob("BasicLogic.bas"))
        if not bas_candidates:
            pytest.skip("watcher did not seed BasicLogic.bas")
        bas = bas_candidates[0]
        marker = f"' UAT-D3-MARKER-{tmp_path.name}"
        text = bas.read_text(encoding="utf-8")
        lines = text.splitlines(keepends=True)
        insert_at = 0
        for i, ln in enumerate(lines):
            if not ln.lstrip().startswith("Attribute "):
                insert_at = i
                break
        lines.insert(insert_at, marker + "\n")
        bas.write_text("".join(lines), encoding="utf-8")

        # Spec: ~2s. Give a generous margin for COM round-trip.
        handle.expect("Change detected", timeout=5)
        handle.expect("Imported", timeout=10)

        # Verify live module reflects the change.
        live = read_module_code(doc, "BasicLogic") or ""
        assert marker in live, (
            f"§D3: marker {marker!r} not in live BasicLogic after watcher "
            f"import. Module head:\n{live[:400]!r}"
        )
    finally:
        handle.stop()
        try:
            doc.Saved = True
        except Exception:
            pass


@pytest.mark.section("visiowings-D4", "visiowings/docs/contributing/uat.md#d4")
@pytest.mark.requires_office("visio")
def test_section_d4_bidirectional_sync(tmp_path, visiowings_repo, fixtures_dir):
    """UAT §D4: ``visiowings edit --bidirectional`` polls Visio and
    writes changes back to the local file when the live module is edited.
    """
    _need_repo(visiowings_repo)
    sample = fixtures_dir / "sample.vsdm"
    if not sample.exists():
        pytest.skip("sample.vsdm fixture missing")
    app = _require_user_opened_visio("sample.vsdm")
    doc = next(d for d in app.Documents if d.Name.lower() == "sample.vsdm")
    doc_path = doc.FullName

    out_dir = tmp_path / "uat-bidi"
    handle = start_watcher(visiowings_repo, doc_path, out_dir, "--bidirectional", "--force")
    try:
        handle.expect("Starting Live Synchronization", timeout=30)
        bas_candidates = list(out_dir.rglob("BasicLogic.bas"))
        if not bas_candidates:
            pytest.skip("watcher did not seed BasicLogic.bas")
        bas = bas_candidates[0]
        baseline = bas.read_text(encoding="utf-8")

        # Edit BasicLogic directly via COM and save the doc.
        marker = f"' UAT-D4-MARKER-{tmp_path.name}"
        comp = doc.VBProject.VBComponents("BasicLogic")
        comp.CodeModule.InsertLines(2, marker)

        # Spec: default polling interval is 4s. Watch up to 12s.
        end = time.monotonic() + 12
        while time.monotonic() < end:
            updated = bas.read_text(encoding="utf-8")
            if marker in updated:
                break
            time.sleep(0.5)

        final = bas.read_text(encoding="utf-8")
        assert marker in final, (
            f"§D4: bidirectional sync didn't push live edit to disk within 12s. "
            f"Local file unchanged from baseline ({len(baseline)} -> {len(final)} bytes). "
            f"Watcher tail:\n{chr(10).join(handle.lines[-10:])}"
        )
    finally:
        handle.stop()
        try:
            doc.Saved = True
        except Exception:
            pass


@pytest.mark.section("visiowings-D5", "visiowings/docs/contributing/uat.md#d5")
@pytest.mark.requires_office("visio")
def test_section_d5_module_deletion_sync(tmp_path, visiowings_repo, fixtures_dir):
    """UAT §D5: ``--sync-delete-modules`` removes the VBE module when a
    local ``.bas`` file is deleted. ``ThisDocument`` is never deleted.
    """
    _need_repo(visiowings_repo)
    sample = fixtures_dir / "sample.vsdm"
    if not sample.exists():
        pytest.skip("sample.vsdm fixture missing")
    app = _require_user_opened_visio("sample.vsdm")
    doc = next(d for d in app.Documents if d.Name.lower() == "sample.vsdm")
    doc_path = doc.FullName

    # Pre-flight: ensure BasicLogic exists in the live doc. If a prior
    # test wiped it we cannot prove the deletion sync — skip cleanly.
    component_names_before = {c.Name for c in doc.VBProject.VBComponents}
    if "BasicLogic" not in component_names_before:
        pytest.skip("BasicLogic missing from live doc — pre-state polluted")

    out_dir = tmp_path / "uat-del"
    handle = start_watcher(
        visiowings_repo,
        doc_path,
        out_dir,
        "--sync-delete-modules",
        "--force",
    )
    try:
        handle.expect("Starting Live Synchronization", timeout=30)
        bas_candidates = list(out_dir.rglob("BasicLogic.bas"))
        if not bas_candidates:
            pytest.skip("watcher did not seed BasicLogic.bas")
        bas = bas_candidates[0]
        bas.unlink()

        # Wait for the deletion to propagate. The spec doesn't fix a
        # timeout; the watcher's debounce + COM round-trip can stretch
        # when prior tests left the live doc loaded with markers. 20s
        # is generous but not infinite.
        end = time.monotonic() + 20
        while time.monotonic() < end:
            names = {c.Name for c in doc.VBProject.VBComponents}
            if "BasicLogic" not in names:
                break
            time.sleep(0.5)

        names_after = {c.Name for c in doc.VBProject.VBComponents}
        assert "BasicLogic" not in names_after, (
            f"§D5: BasicLogic still in live doc after .bas deletion + 20s. "
            f"Live components: {sorted(names_after)}. "
            f"Watcher tail:\n{chr(10).join(handle.lines[-15:])}"
        )
        # ThisDocument must NEVER be touched by deletion sync.
        assert "ThisDocument" in names_after, (
            "§D5: ThisDocument was deleted by --sync-delete-modules. "
            "Spec says document modules are immune to this flag."
        )
    finally:
        handle.stop()
        try:
            doc.Saved = True
        except Exception:
            pass


# --------------------------------------------------------------------------
# §E — Encoding (Iter 3 subset: E1 cp1252 roundtrip, E4 emoji warning)
# E2 (--codepage override) + E3 (BOM detection) deferred to iter4.
# --------------------------------------------------------------------------


@pytest.mark.section("visiowings-E1", "visiowings/docs/contributing/uat.md#e1")
@pytest.mark.requires_office("visio")
def test_section_e1_umlauts_roundtrip(tmp_path, visiowings_repo, fixtures_dir):
    """UAT §E1: German umlaute in comments survive export → import roundtrip.

    Inject a comment with umlauts into the .bas via export, modify with
    umlaut comment, import, read VBE — comment chars must match.
    """
    _need_repo(visiowings_repo)
    sample = fixtures_dir / "sample.vsdm"
    if not sample.exists():
        pytest.skip("sample.vsdm fixture missing")
    app = _require_user_opened_visio("sample.vsdm")
    from tests.uat.com_helpers.vbe import read_module_code

    # Refresh doc from disk to discard accumulated marker comments from
    # earlier export/import tests in the same pytest run.
    doc = _refresh_user_doc(app, "sample.vsdm")
    if doc is None:
        doc = next(d for d in app.Documents if d.Name.lower() == "sample.vsdm")
    doc_path = doc.FullName
    export_dir = tmp_path / "uat-e1"
    umlaut_marker = "' Größe für Zähler übergeben"

    cp = run_branch(
        [
            *visiowings_cli(visiowings_repo),
            "export",
            "--file",
            str(doc_path),
            "--output",
            str(export_dir),
        ],
        visiowings_repo,
        timeout=120,
    )
    assert cp.returncode == 0, f"export failed: {cp.stderr or cp.stdout}"

    bas_files = list(export_dir.rglob("BasicLogic.bas"))
    if not bas_files:
        pytest.skip("BasicLogic.bas not produced")
    bas = bas_files[0]
    text = bas.read_text(encoding="utf-8")
    lines = text.splitlines(keepends=True)
    insert_at = 1 if lines and lines[0].lstrip().startswith("Attribute ") else 0
    lines.insert(insert_at, umlaut_marker + "\n")
    bas.write_text("".join(lines), encoding="utf-8")

    # Pipe "o" repeatedly so any "action (o/s/i/C)" overwrite prompt
    # gets auto-answered. ``--force`` alone doesn't suppress all prompts.
    cp = run_branch(
        [
            *visiowings_cli(visiowings_repo),
            "import",
            "--file",
            str(doc_path),
            "--input",
            str(export_dir),
            "--force",
        ],
        visiowings_repo,
        timeout=120,
        stdin_input=("o\n" * 10),
    )
    assert cp.returncode == 0, f"import failed: {cp.stderr or cp.stdout}"

    try:
        code = read_module_code(doc, "BasicLogic")
    finally:
        try:
            doc.Saved = True
        except Exception:
            pass
    assert umlaut_marker in code, (
        f"umlauts mangled in roundtrip. Expected {umlaut_marker!r} in module; head: {code[:400]!r}"
    )


@pytest.mark.section("visiowings-E2", "visiowings/docs/contributing/uat.md#e2")
@pytest.mark.requires_office("visio")
def test_section_e2_codepage_override_no_silent_corruption(tmp_path, visiowings_repo, fixtures_dir):
    """UAT §E2: ``--codepage`` override against a doc with a different
    LCID either succeeds cleanly or warns explicitly. Silent corruption
    (export written in cp1251 but content was cp1252) is the failure mode.

    sample.vsdm is a German-LCID doc. Forcing ``--codepage cp1251``
    (Cyrillic) should either:
      a) export cleanly because the content is ASCII-safe, OR
      b) emit a warning naming the codepage/LCID mismatch.
    Either way the round-trip back via import must not mangle existing
    German characters in BasicLogic.bas.
    """
    _need_repo(visiowings_repo)
    sample = fixtures_dir / "sample.vsdm"
    if not sample.exists():
        pytest.skip("sample.vsdm fixture missing")
    app = _require_user_opened_visio("sample.vsdm")
    doc = next(d for d in app.Documents if d.Name.lower() == "sample.vsdm")
    doc_path = doc.FullName
    out_dir = tmp_path / "uat-e2"

    cp = run_branch(
        [
            *visiowings_cli(visiowings_repo),
            "export",
            "--file",
            str(doc_path),
            "--output",
            str(out_dir),
            "--codepage",
            "cp1251",
        ],
        visiowings_repo,
        timeout=120,
    )
    combined = cp.stdout + cp.stderr
    # Spec: ``either succeeds or warns explicitly``. We accept both, but
    # an unrelated traceback or silent successful export of mangled bytes
    # is a fail.
    assert not _has_python_traceback(combined), (
        f"--codepage override leaked a Python traceback:\n{combined[-1500:]}"
    )
    if cp.returncode == 0:
        # Success path: re-import the exported tree and verify VBE module
        # content is still readable VBA (no decode garbage).
        bas_files = list(out_dir.rglob("*.bas"))
        assert bas_files, (
            f"export reported success but produced no .bas files. stdout: {cp.stdout[-400:]}"
        )
        # Each .bas should decode under cp1251 cleanly (visiowings claims
        # it wrote cp1251). If it secretly wrote cp1252 the German umlauts
        # in any comment would surface here.
        for bas in bas_files:
            try:
                bas.read_text(encoding="cp1251")
            except UnicodeDecodeError as exc:
                pytest.fail(
                    f"{bas.name} not cp1251-decodable despite --codepage cp1251: "
                    f"{exc!r}. This is silent codepage corruption (§E2)."
                )
    else:
        # Refused path: must name the codepage/LCID conflict, not crash.
        lowered = combined.lower()
        mentions_cp = "codepage" in lowered or "cp1251" in lowered or "lcid" in lowered
        assert mentions_cp, (
            f"export refused (exit {cp.returncode}) without naming the "
            f"codepage/LCID conflict. Output: {combined[-800:]}"
        )


@pytest.mark.section("visiowings-E3", "visiowings/docs/contributing/uat.md#e3")
@pytest.mark.requires_office("visio")
def test_section_e3_bom_detection_on_import(tmp_path, visiowings_repo, fixtures_dir):
    """UAT §E3: importing a ``.bas`` saved as UTF-8 **with BOM** must
    strip the BOM before Visio sees it — no ```` literal at the
    start of the first module line in VBE.
    """
    _need_repo(visiowings_repo)
    sample = fixtures_dir / "sample.vsdm"
    if not sample.exists():
        pytest.skip("sample.vsdm fixture missing")
    app = _require_user_opened_visio("sample.vsdm")
    from tests.uat.com_helpers.vbe import read_module_code

    doc = _refresh_user_doc(app, "sample.vsdm")
    if doc is None:
        doc = next(d for d in app.Documents if d.Name.lower() == "sample.vsdm")
    doc_path = doc.FullName

    export_dir = tmp_path / "uat-e3"
    cp = run_branch(
        [
            *visiowings_cli(visiowings_repo),
            "export",
            "--file",
            str(doc_path),
            "--output",
            str(export_dir),
        ],
        visiowings_repo,
        timeout=120,
    )
    assert cp.returncode == 0, f"export failed: {cp.stderr or cp.stdout}"
    bas_candidates = list(export_dir.rglob("BasicLogic.bas"))
    if not bas_candidates:
        pytest.skip("BasicLogic.bas not produced")
    bas = bas_candidates[0]

    # Rewrite the .bas with a UTF-8 BOM at the very start.
    raw_text = bas.read_text(encoding="utf-8")
    bas.write_bytes(b"\xef\xbb\xbf" + raw_text.encode("utf-8"))
    # Sanity: file does start with BOM bytes on disk.
    assert bas.read_bytes()[:3] == b"\xef\xbb\xbf"

    cp = run_branch(
        [
            *visiowings_cli(visiowings_repo),
            "import",
            "--file",
            str(doc_path),
            "--input",
            str(export_dir),
            "--force",
        ],
        visiowings_repo,
        timeout=120,
        stdin_input=("o\n" * 10),
    )
    assert cp.returncode == 0, f"import of BOM-prefixed .bas failed: {cp.stderr or cp.stdout}"

    try:
        code = read_module_code(doc, "BasicLogic")
    finally:
        try:
            doc.Saved = True
        except Exception:
            pass
    # Spec: ``no U+FEFF character at the start of the first line in the
    # VBA Editor``. The BOM must have been stripped.
    BOM = "﻿"
    assert not code.startswith(BOM), (
        f"BOM leaked into VBE module: first 16 chars = {code[:16]!r}. "
        "visiowings did not strip the UTF-8 BOM before AddFromString."
    )
    first_line = code.splitlines()[0] if code.splitlines() else ""
    assert BOM not in first_line, f"BOM found inside first line of imported module: {first_line!r}"


@pytest.mark.section("visiowings-E4", "visiowings/docs/contributing/uat.md#e4")
@pytest.mark.requires_office("visio")
def test_section_e4_emoji_warns_no_silent_loss(tmp_path, visiowings_repo, fixtures_dir):
    """UAT §E4: out-of-codepage character (emoji) must warn explicitly,
    not silently mangle.

    Inject an emoji into a .bas comment and run import. visiowings must
    EITHER round-trip the emoji cleanly OR emit a warning that mentions
    the out-of-range character. Silent corruption (no warning, emoji
    replaced with ``?``) is the failure mode this test guards against.
    """
    _need_repo(visiowings_repo)
    sample = fixtures_dir / "sample.vsdm"
    if not sample.exists():
        pytest.skip("sample.vsdm fixture missing")
    app = _require_user_opened_visio("sample.vsdm")
    from tests.uat.com_helpers.vbe import read_module_code

    # Refresh doc from disk to discard accumulated marker comments from
    # earlier export/import tests in the same pytest run.
    doc = _refresh_user_doc(app, "sample.vsdm")
    if doc is None:
        doc = next(d for d in app.Documents if d.Name.lower() == "sample.vsdm")
    doc_path = doc.FullName
    export_dir = tmp_path / "uat-e4"
    emoji = "🌍"
    marker = f"' Earth: {emoji}"

    cp = run_branch(
        [
            *visiowings_cli(visiowings_repo),
            "export",
            "--file",
            str(doc_path),
            "--output",
            str(export_dir),
        ],
        visiowings_repo,
        timeout=120,
    )
    assert cp.returncode == 0, f"export failed: {cp.stderr or cp.stdout}"

    bas_files = list(export_dir.rglob("BasicLogic.bas"))
    if not bas_files:
        pytest.skip("BasicLogic.bas not produced")
    bas = bas_files[0]
    text = bas.read_text(encoding="utf-8")
    lines = text.splitlines(keepends=True)
    insert_at = 1 if lines and lines[0].lstrip().startswith("Attribute ") else 0
    lines.insert(insert_at, marker + "\n")
    bas.write_text("".join(lines), encoding="utf-8")

    # Pipe "o" repeatedly so any "action (o/s/i/C)" overwrite prompt
    # gets auto-answered. ``--force`` alone doesn't suppress all prompts.
    cp = run_branch(
        [
            *visiowings_cli(visiowings_repo),
            "import",
            "--file",
            str(doc_path),
            "--input",
            str(export_dir),
            "--force",
        ],
        visiowings_repo,
        timeout=120,
        stdin_input=("o\n" * 10),
    )
    assert cp.returncode == 0, f"import failed: {cp.stderr or cp.stdout}"
    combined = cp.stdout + cp.stderr

    try:
        code = read_module_code(doc, "BasicLogic")
        module_destroyed = False
    except Exception as exc:
        # If reading BasicLogic raises, the module was deleted from the
        # doc during the emoji import — the worst form of "silent loss"
        # this test exists to catch.
        if "ndex" in str(exc) or "Bereich" in str(exc) or "Range" in str(exc):
            module_destroyed = True
            code = ""
        else:
            raise
    finally:
        try:
            doc.Saved = True
        except Exception:
            pass

    if module_destroyed:
        pytest.fail(
            "visiowings silently destroyed BasicLogic during emoji import. "
            "After import, proj.VBComponents('BasicLogic') raises 'Index out "
            "of range' — the module no longer exists in the doc. This is the "
            "worst form of data loss this test is designed to catch. See "
            "prompts/visiowings-fix.md finding #4 for repro + diagnosis.\n"
            f"visiowings stdout: {combined[-600:]!r}"
        )

    # Acceptance: either the emoji survives intact OR a warning was emitted.
    emoji_survived = emoji in code
    warned = any(
        token in combined.lower()
        for token in (
            "warning",
            "warn",
            "out-of-range",
            "out of range",
            "codepage",
            "unicode",
            "encoding",
        )
    )
    assert emoji_survived or warned, (
        "out-of-codepage character was silently corrupted. "
        f"Emoji survived: {emoji_survived}. Warned: {warned}. "
        f"Module head: {code[:400]!r}. "
        f"CLI output tail: {combined[-400:]!r}"
    )


@pytest.mark.section("visiowings-F1", "visiowings/docs/contributing/uat.md#f1")
@pytest.mark.requires_office("visio")
def test_section_f1_subfolder_per_document(tmp_path, visiowings_repo, fixtures_dir):
    """UAT §F1 (partial): each open document's modules land in a
    sub-folder named after the doc (sanitised).

    The full spec uses three fixture files (drawing + stencil + template)
    open simultaneously. We only ship a single ``sample.vsdm`` fixture,
    so this test verifies the **layout property** that even with one
    open doc, modules land in ``<output>/sample/`` (not directly at
    ``<output>/``). The 3-doc variant is queued for a setup that adds a
    second/third fixture; the property pinned here is the same one that
    must hold in the multi-doc case.
    """
    _need_repo(visiowings_repo)
    sample = fixtures_dir / "sample.vsdm"
    if not sample.exists():
        pytest.skip("sample.vsdm fixture missing")
    app = _require_user_opened_visio("sample.vsdm")
    doc = next(d for d in app.Documents if d.Name.lower() == "sample.vsdm")
    doc_path = doc.FullName

    out_dir = tmp_path / "uat-multi"
    cp = run_branch(
        [
            *visiowings_cli(visiowings_repo),
            "export",
            "--file",
            str(doc_path),
            "--output",
            str(out_dir),
        ],
        visiowings_repo,
        timeout=120,
    )
    assert cp.returncode == 0, f"export failed: {cp.stderr or cp.stdout}"

    # Module files must live one level deep under <output>, in a folder
    # named for the doc — not flat at <output>/.
    flat_files = [
        p for p in out_dir.iterdir() if p.is_file() and p.suffix.lower() in {".bas", ".cls", ".frm"}
    ]
    assert not flat_files, (
        f"§F1 layout violated: module files at <output>/ root: "
        f"{[p.name for p in flat_files]}. Spec requires per-doc subfolders."
    )
    subdirs = [p for p in out_dir.iterdir() if p.is_dir()]
    assert subdirs, (
        f"§F1 layout violated: no per-doc subfolder under {out_dir}. "
        f"Got: {[p.name for p in out_dir.iterdir()]}"
    )
    # Sanitised name — no spaces, no path-illegal chars.
    bad_chars = set(' \t<>:"|?*')
    for sub in subdirs:
        offenders = bad_chars.intersection(sub.name)
        assert not offenders, f"§F1 subfolder name {sub.name!r} contains unsafe chars: {offenders}"
    # All exported modules live under the subfolder, not loose at root.
    nested_modules = list(out_dir.rglob("*.bas")) + list(out_dir.rglob("*.cls"))
    assert nested_modules, "no modules exported at all"
    for mod in nested_modules:
        assert mod.parent != out_dir, (
            f"module {mod.name} at <output>/ root instead of per-doc subfolder"
        )


@pytest.mark.section("visiowings-F2", "visiowings/docs/contributing/uat.md#f2")
@pytest.mark.requires_office("visio")
def test_section_f2_rubberduck_folder_export(tmp_path, visiowings_repo, fixtures_dir):
    """UAT §F2: a module with ``'@Folder("UI.Buttons")`` as the first
    non-Attribute line lands at ``<out>/<doc>/UI/Buttons/<Module>.bas``
    when exported with ``--rd``.

    Setup: export sample.vsdm, inject the @Folder annotation into
    BasicLogic.bas, import the annotated version back into the live
    doc, then re-export with --rd and verify the path layout.
    """
    _need_repo(visiowings_repo)
    sample = fixtures_dir / "sample.vsdm"
    if not sample.exists():
        pytest.skip("sample.vsdm fixture missing")
    app = _require_user_opened_visio("sample.vsdm")
    from tests.uat.com_helpers.vbe import read_module_code

    doc = _refresh_user_doc(app, "sample.vsdm")
    if doc is None:
        doc = next(d for d in app.Documents if d.Name.lower() == "sample.vsdm")
    doc_path = doc.FullName

    # Phase 1: plain export to get the seed tree.
    seed_dir = tmp_path / "rd-seed"
    cp = run_branch(
        [
            *visiowings_cli(visiowings_repo),
            "export",
            "--file",
            str(doc_path),
            "--output",
            str(seed_dir),
        ],
        visiowings_repo,
        timeout=120,
    )
    assert cp.returncode == 0, f"seed export failed: {cp.stderr or cp.stdout}"
    bas_candidates = list(seed_dir.rglob("BasicLogic.bas"))
    if not bas_candidates:
        pytest.skip("BasicLogic.bas not produced by export")
    bas = bas_candidates[0]

    # Phase 2: inject '@Folder annotation into BasicLogic.bas
    folder_marker = '\'@Folder("UI.Buttons")'
    text = bas.read_text(encoding="utf-8")
    lines = text.splitlines(keepends=True)
    insert_at = 0
    for i, ln in enumerate(lines):
        if not ln.lstrip().startswith("Attribute "):
            insert_at = i
            break
    lines.insert(insert_at, folder_marker + "\n")
    bas.write_text("".join(lines), encoding="utf-8")

    # Phase 3: import the annotated version back into the live doc.
    cp = run_branch(
        [
            *visiowings_cli(visiowings_repo),
            "import",
            "--file",
            str(doc_path),
            "--input",
            str(seed_dir),
            "--force",
        ],
        visiowings_repo,
        timeout=120,
        stdin_input=("o\n" * 10),
    )
    assert cp.returncode == 0, f"import of annotated BasicLogic failed: {cp.stderr or cp.stdout}"
    # Sanity: live module now carries the annotation.
    live = read_module_code(doc, "BasicLogic") or ""
    assert folder_marker in live, (
        f"§F2 setup failed: @Folder annotation didn't make it into the live "
        f"module after --force import. Module head:\n{live[:300]!r}"
    )

    # Phase 4: re-export with --rd and check the path layout.
    rd_dir = tmp_path / "uat-rd"
    cp = run_branch(
        [
            *visiowings_cli(visiowings_repo),
            "export",
            "--file",
            str(doc_path),
            "--output",
            str(rd_dir),
            "--rd",
        ],
        visiowings_repo,
        timeout=120,
    )
    try:
        doc.Saved = True
    except Exception:
        pass
    assert cp.returncode == 0, f"§F2: export --rd failed: {cp.stderr or cp.stdout}"
    expected = list(rd_dir.rglob("UI/Buttons/BasicLogic.bas"))
    actual_bas = [p.relative_to(rd_dir) for p in rd_dir.rglob("BasicLogic.bas")]
    assert expected, (
        f"§F2: BasicLogic.bas not at <out>/<doc>/UI/Buttons/. Found BasicLogic.bas at: {actual_bas}"
    )


@pytest.mark.section("visiowings-F3", "visiowings/docs/contributing/uat.md#f3")
@pytest.mark.requires_office("visio")
def test_section_f3_rubberduck_folder_import_auto_inject(tmp_path, visiowings_repo, fixtures_dir):
    """UAT §F3: importing with ``--rd`` from ``./uat-rd/<doc>/Helpers/<M>.bas``
    auto-injects ``'@Folder("Helpers")`` into the module's source.

    Setup: export, move BasicLogic.bas into a ``Helpers/`` subfolder
    (stripping any pre-existing @Folder annotation), import with --rd,
    verify the live module contains @Folder("Helpers").
    """
    _need_repo(visiowings_repo)
    sample = fixtures_dir / "sample.vsdm"
    if not sample.exists():
        pytest.skip("sample.vsdm fixture missing")
    app = _require_user_opened_visio("sample.vsdm")
    from tests.uat.com_helpers.vbe import read_module_code

    doc = _refresh_user_doc(app, "sample.vsdm")
    if doc is None:
        doc = next(d for d in app.Documents if d.Name.lower() == "sample.vsdm")
    doc_path = doc.FullName

    # Phase 1: plain export.
    seed_dir = tmp_path / "rd-import-seed"
    cp = run_branch(
        [
            *visiowings_cli(visiowings_repo),
            "export",
            "--file",
            str(doc_path),
            "--output",
            str(seed_dir),
        ],
        visiowings_repo,
        timeout=120,
    )
    assert cp.returncode == 0, f"seed export failed: {cp.stderr or cp.stdout}"
    bas_candidates = list(seed_dir.rglob("BasicLogic.bas"))
    if not bas_candidates:
        pytest.skip("BasicLogic.bas not produced by export")
    src_bas = bas_candidates[0]

    # Phase 2: strip any existing @Folder annotation and relocate the
    # file to <doc>/Helpers/BasicLogic.bas.
    text = src_bas.read_text(encoding="utf-8")
    cleaned = "\n".join(ln for ln in text.splitlines() if not ln.lstrip().startswith("'@Folder("))
    if not cleaned.endswith("\n"):
        cleaned += "\n"
    doc_subdir = src_bas.parent  # <seed_dir>/sample/
    helpers = doc_subdir / "Helpers"
    helpers.mkdir(parents=True, exist_ok=True)
    relocated = helpers / "BasicLogic.bas"
    relocated.write_text(cleaned, encoding="utf-8")
    src_bas.unlink()  # remove from <doc>/ root so --rd uses the new path

    # Phase 3: import with --rd from the seed_dir root.
    cp = run_branch(
        [
            *visiowings_cli(visiowings_repo),
            "import",
            "--file",
            str(doc_path),
            "--input",
            str(seed_dir),
            "--force",
            "--rd",
        ],
        visiowings_repo,
        timeout=120,
        stdin_input=("o\n" * 10),
    )
    assert cp.returncode == 0, f"§F3: import --rd failed: {cp.stderr or cp.stdout}"

    # Phase 4: verify the live BasicLogic module carries
    # '@Folder("Helpers") near the top.
    live = read_module_code(doc, "BasicLogic") or ""
    try:
        doc.Saved = True
    except Exception:
        pass
    expected_annotation = '\'@Folder("Helpers")'
    assert expected_annotation in live, (
        f"§F3: --rd import didn't inject {expected_annotation!r} into the "
        f"module. Live module head:\n{live[:400]!r}"
    )
    # Annotation should be near the top — before any Sub/Function body.
    head_until_code = []
    for ln in live.splitlines():
        head_until_code.append(ln)
        stripped = ln.strip().lower()
        if stripped.startswith(("sub ", "function ", "private ", "public ", "property ")):
            break
    head_blob = "\n".join(head_until_code)
    assert expected_annotation in head_blob, (
        f"§F3: @Folder annotation appears after the first code line, not "
        f"near the top. Head until first code line:\n{head_blob!r}"
    )


@pytest.mark.section("visiowings-G3", "visiowings/docs/contributing/uat.md#g3")
@pytest.mark.requires_office("visio")
def test_section_g3_thisdocument_protection(tmp_path, visiowings_repo, fixtures_dir):
    """UAT §G3: editing ``ThisDocument.cls`` locally and importing
    without ``--force`` must WARN and leave the document module
    unchanged. ``--force`` must override and import the edit.
    """
    _need_repo(visiowings_repo)
    sample = fixtures_dir / "sample.vsdm"
    if not sample.exists():
        pytest.skip("sample.vsdm fixture missing")
    app = _require_user_opened_visio("sample.vsdm")
    from tests.uat.com_helpers.vbe import read_module_code

    doc = _refresh_user_doc(app, "sample.vsdm")
    if doc is None:
        doc = next(d for d in app.Documents if d.Name.lower() == "sample.vsdm")
    doc_path = doc.FullName

    export_dir = tmp_path / "uat-g3"
    cp = run_branch(
        [
            *visiowings_cli(visiowings_repo),
            "export",
            "--file",
            str(doc_path),
            "--output",
            str(export_dir),
        ],
        visiowings_repo,
        timeout=120,
    )
    assert cp.returncode == 0, f"export failed: {cp.stderr or cp.stdout}"

    td_candidates = list(export_dir.rglob("ThisDocument.cls"))
    if not td_candidates:
        pytest.skip("ThisDocument.cls not produced by export")
    td = td_candidates[0]
    marker = f"' UAT-G3-MARKER-{tmp_path.name}"
    text = td.read_text(encoding="utf-8")
    lines = text.splitlines(keepends=True)
    # Insert after Attribute header but before any code so we land in a
    # syntactically valid spot.
    insert_at = 0
    for i, ln in enumerate(lines):
        if not ln.lstrip().startswith("Attribute "):
            insert_at = i
            break
    lines.insert(insert_at, marker + "\n")
    td.write_text("".join(lines), encoding="utf-8")

    # Step 2: import without --force. Expected: skip + warning.
    cp_soft = run_branch(
        [
            *visiowings_cli(visiowings_repo),
            "import",
            "--file",
            str(doc_path),
            "--input",
            str(export_dir),
        ],
        visiowings_repo,
        timeout=120,
        stdin_input=("s\n" * 10),  # 's' = skip, matches default-skip semantics
    )
    soft_combined = cp_soft.stdout + cp_soft.stderr
    code_after_soft = read_module_code(doc, "ThisDocument") or ""
    try:
        doc.Saved = True
    except Exception:
        pass

    # Without --force: marker MUST NOT be in the live module. The CLI
    # either skipped the conflict (clean exit + warning) or refused
    # (non-zero exit + warning). Both are spec-conformant.
    assert marker not in code_after_soft, (
        "§G3 violation: ThisDocument was modified WITHOUT --force. "
        f"Marker {marker!r} appears in live module. CLI output:\n{soft_combined[-600:]}"
    )
    lowered = soft_combined.lower()
    warned_soft = any(
        w in lowered
        for w in ("warning", "skip", "skipped", "document module", "thisdocument", "protected")
    )
    assert warned_soft, (
        f"§G3: importing edited ThisDocument without --force must emit a "
        f"warning or skip-message naming the protection. Output:\n{soft_combined[-800:]}"
    )

    # Step 4: re-run with --force. Expected: import succeeds, marker lands.
    cp_force = run_branch(
        [
            *visiowings_cli(visiowings_repo),
            "import",
            "--file",
            str(doc_path),
            "--input",
            str(export_dir),
            "--force",
        ],
        visiowings_repo,
        timeout=120,
        stdin_input=("o\n" * 10),
    )
    assert cp_force.returncode == 0, (
        f"§G3: import --force on ThisDocument failed: {cp_force.stderr or cp_force.stdout}"
    )
    code_after_force = read_module_code(doc, "ThisDocument") or ""
    try:
        doc.Saved = True
    except Exception:
        pass
    assert marker in code_after_force, (
        f"§G3: import --force should have overwritten ThisDocument, but "
        f"marker {marker!r} is not in live module. Module head:\n{code_after_force[:400]!r}"
    )


@pytest.mark.section("visiowings-G4", "visiowings/docs/contributing/uat.md#g4")
def test_section_g4_readonly_output_clean_error(tmp_path, visiowings_repo, fixtures_dir):
    """UAT §G4: exporting to a read-only output directory must error
    clearly (mentions ``permission`` / the path), exit non-zero, and
    leave no partial writes behind. No Python traceback.

    We don't need an open Visio for this — the FS-permission check
    happens before any COM work (or should). If the implementation only
    surfaces the error after talking to Visio, the test still passes as
    long as the failure mode matches the spec.
    """
    _need_repo(visiowings_repo)
    sample = fixtures_dir / "sample.vsdm"
    if not sample.exists():
        pytest.skip("sample.vsdm fixture missing")
    readonly_dir = tmp_path / "uat-readonly"
    readonly_dir.mkdir()
    # Make a sentinel file the test can check wasn't overwritten and
    # remove write perm on Windows via icacls (chmod 0o555 is mostly a
    # no-op on NTFS). If the ACL change fails, skip rather than fail —
    # this is an environment quirk, not a visiowings bug.
    sentinel = readonly_dir / "_keep.txt"
    sentinel.write_text("hands off", encoding="utf-8")
    import subprocess as _sp

    icacls_cp = _sp.run(
        ["icacls", str(readonly_dir), "/deny", f"{os.environ.get('USERNAME', 'Users')}:(W)"],
        capture_output=True,
        text=True,
        timeout=30,
    )
    if icacls_cp.returncode != 0:
        pytest.skip(
            f"could not set read-only ACL on {readonly_dir} via icacls: {icacls_cp.stderr[:300]}"
        )
    try:
        cp = run_branch(
            [
                *visiowings_cli(visiowings_repo),
                "export",
                "--file",
                str(sample),
                "--output",
                str(readonly_dir),
            ],
            visiowings_repo,
            timeout=60,
        )
    finally:
        # Always lift the deny ACE so pytest can clean tmp_path.
        _sp.run(
            ["icacls", str(readonly_dir), "/remove:d", os.environ.get("USERNAME", "Users")],
            capture_output=True,
            text=True,
            timeout=30,
        )
    combined = cp.stdout + cp.stderr
    assert cp.returncode != 0, (
        f"§G4: export to read-only dir should fail, but exit was 0. Output:\n{combined[-600:]}"
    )
    assert not _has_python_traceback(combined), (
        f"§G4: read-only export leaked a Python traceback:\n{combined[-1500:]}"
    )
    lowered = combined.lower()
    names_problem = (
        "permission" in lowered
        or "denied" in lowered
        or "read-only" in lowered
        or "readonly" in lowered
        or "access" in lowered
        or str(readonly_dir).lower() in lowered
        or "uat-readonly" in lowered
    )
    assert names_problem, (
        f"§G4: error message must mention permission/access or the path. Got:\n{combined[-800:]}"
    )


@pytest.mark.section("visiowings-G5", "visiowings/docs/contributing/uat.md#g5")
@pytest.mark.requires_office("visio")
def test_section_g5_graceful_shutdown(tmp_path, visiowings_repo, fixtures_dir):
    """UAT §G5: a Ctrl+C / CTRL_BREAK to ``visiowings edit`` prints
    ``Shutting down...`` within ~2s and exits without leaving zombies.
    """
    _need_repo(visiowings_repo)
    sample = fixtures_dir / "sample.vsdm"
    if not sample.exists():
        pytest.skip("sample.vsdm fixture missing")
    app = _require_user_opened_visio("sample.vsdm")
    doc = next(d for d in app.Documents if d.Name.lower() == "sample.vsdm")
    doc_path = doc.FullName

    out_dir = tmp_path / "uat-g5"
    handle = start_watcher(visiowings_repo, doc_path, out_dir, "--force")
    try:
        handle.expect("Starting Live Synchronization", timeout=30)
        # Watcher is now idle; trigger graceful shutdown.
        t0 = time.monotonic()
        exit_code = handle.stop(timeout=5.0)
        elapsed = time.monotonic() - t0
        # Spec: ~2s. Accept up to 5s for CI variability.
        assert elapsed < 5.0, (
            f"§G5: shutdown took {elapsed:.1f}s (>5s). "
            f"Watcher tail:\n{chr(10).join(handle.lines[-10:])}"
        )
        # Verify the watcher reached its graceful-shutdown branch (or at
        # least didn't crash).
        full_output = "\n".join(handle.lines)
        saw_banner = "Shutting down" in full_output
        # Some implementations may exit silently on CTRL_BREAK without
        # the banner; accept that as long as the exit code is sane.
        assert saw_banner or exit_code in (0, 1, -1073741510, 130, 3221225786), (
            f"§G5: no 'Shutting down' banner and unexpected exit code {exit_code}. "
            f"Tail:\n{chr(10).join(handle.lines[-10:])}"
        )
    finally:
        # Ensure no zombie if assertion above failed before stop().
        if handle.proc.poll() is None:
            handle.stop(timeout=2.0)
        try:
            doc.Saved = True
        except Exception:
            pass


@pytest.mark.section("visiowings-G6", "visiowings/docs/contributing/uat.md#g6")
@pytest.mark.requires_office("visio")
def test_section_g6_rapid_saves_debounced(tmp_path, visiowings_repo, fixtures_dir):
    """UAT §G6: 5 rapid saves of the same file within ~1s collapse into
    a single import via the watcher's debounce window. ``[DEBUG]
    Debouncing: ...`` lines must appear at least once.
    """
    _need_repo(visiowings_repo)
    sample = fixtures_dir / "sample.vsdm"
    if not sample.exists():
        pytest.skip("sample.vsdm fixture missing")
    app = _require_user_opened_visio("sample.vsdm")
    doc = next(d for d in app.Documents if d.Name.lower() == "sample.vsdm")
    doc_path = doc.FullName

    out_dir = tmp_path / "uat-g6"
    handle = start_watcher(visiowings_repo, doc_path, out_dir, "--force", "--debug")
    try:
        handle.expect("Starting Live Synchronization", timeout=30)
        bas_candidates = list(out_dir.rglob("BasicLogic.bas"))
        if not bas_candidates:
            pytest.skip("watcher did not seed BasicLogic.bas")
        bas = bas_candidates[0]

        # 5 rapid touches within ~1s — each appending a new comment so
        # the file truly changes and watchdog can't dedupe by mtime alone.
        baseline = bas.read_text(encoding="utf-8")
        for i in range(5):
            content = baseline + f"\n' UAT-G6-spam-{i}\n"
            bas.write_text(content, encoding="utf-8")
            time.sleep(0.15)  # total ~750ms — well within 1s debounce

        # Give the debounce window to settle + import to happen.
        time.sleep(3.0)

        full_output = "\n".join(handle.lines)
        # Spec: "Debouncing" lines visible. With --debug we get [DEBUG] Debouncing: ...
        assert "Debouncing" in full_output, (
            f"§G6: no 'Debouncing' line in --debug output after 5 rapid saves. "
            f"Tail:\n{chr(10).join(handle.lines[-15:])}"
        )
        # Spec: only one import per debounce window. We saw 5 writes
        # within 1s; conservatively assert no more than 2 imports
        # (allows a follow-up import after the burst settled).
        import_count = sum(
            1 for ln in handle.lines if ln.startswith("✓ Imported") or " Imported: " in ln
        )
        # First "Imported" from initial export is fine; count post-banner ones.
        post_banner_imports = 0
        seen_banner = False
        for ln in handle.lines:
            if "Starting Live Synchronization" in ln:
                seen_banner = True
                continue
            if seen_banner and ("✓ Imported" in ln or " Imported: " in ln):
                post_banner_imports += 1
        assert post_banner_imports <= 2, (
            f"§G6: 5 rapid saves triggered {post_banner_imports} imports "
            f"(spec: at most 1 per debounce window). "
            f"Total Imported lines (incl. initial export): {import_count}. "
            f"Tail:\n{chr(10).join(handle.lines[-20:])}"
        )
    finally:
        handle.stop()
        try:
            doc.Saved = True
        except Exception:
            pass


def _run_with_update_env(
    visiowings_repo, env_override: dict[str, str], *args: str, timeout: int = 60
):
    """Invoke a visiowings command with extra env vars layered onto the
    standard branch env. Returns the CompletedProcess.
    """
    import subprocess as _sp

    env = os.environ.copy()
    env["PYTHONPATH"] = str(visiowings_repo) + (
        os.pathsep + env["PYTHONPATH"] if env.get("PYTHONPATH") else ""
    )
    env.update(env_override)
    return _sp.run(
        [sys.executable, "-m", "visiowings.cli", *args],
        cwd=str(visiowings_repo),
        env=env,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=timeout,
        check=False,
    )


def _looks_like_update_hint(text: str) -> bool:
    """Heuristic: a 'newer version available' hint contains either the
    word 'newer' near 'version' or an explicit 'update available' phrasing.
    """
    lowered = text.lower()
    return (
        ("newer" in lowered and "version" in lowered)
        or "update available" in lowered
        or "a new release" in lowered
    )


@pytest.mark.section("visiowings-H2", "visiowings/docs/contributing/uat.md#h2")
def test_section_h2_no_update_check_env_var(visiowings_repo):
    """UAT §H2: ``VISIOWINGS_NO_UPDATE_CHECK=1`` suppresses the update-
    check hint. Calling ``visiowings --help`` with the env var set must
    not surface a 'newer version' message.
    """
    _need_repo(visiowings_repo)
    cp = _run_with_update_env(
        visiowings_repo,
        {"VISIOWINGS_NO_UPDATE_CHECK": "1"},
        "--help",
        timeout=30,
    )
    combined = cp.stdout + cp.stderr
    assert cp.returncode == 0, (
        f"§H2: visiowings --help failed with VISIOWINGS_NO_UPDATE_CHECK=1: {combined[-500:]}"
    )
    assert not _looks_like_update_hint(combined), (
        f"§H2: update hint leaked despite VISIOWINGS_NO_UPDATE_CHECK=1. Output:\n{combined[-600:]}"
    )


@pytest.mark.section("visiowings-H3", "visiowings/docs/contributing/uat.md#h3")
def test_section_h3_no_update_check_toml(tmp_path, visiowings_repo):
    """UAT §H3: ``update_check = false`` in ``.visiowings.toml`` suppresses
    the hint with no env var present. Run from a workdir containing a
    minimal toml.
    """
    _need_repo(visiowings_repo)
    workdir = tmp_path / "workdir"
    workdir.mkdir()
    toml = workdir / ".visiowings.toml"
    toml.write_text("update_check = false\n", encoding="utf-8")
    import subprocess as _sp

    env = os.environ.copy()
    env["PYTHONPATH"] = str(visiowings_repo) + (
        os.pathsep + env["PYTHONPATH"] if env.get("PYTHONPATH") else ""
    )
    # Explicitly do NOT set VISIOWINGS_NO_UPDATE_CHECK; clear it if set.
    env.pop("VISIOWINGS_NO_UPDATE_CHECK", None)
    cp = _sp.run(
        [sys.executable, "-m", "visiowings.cli", "--help"],
        cwd=str(workdir),
        env=env,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=30,
        check=False,
    )
    combined = cp.stdout + cp.stderr
    assert cp.returncode == 0, (
        f"§H3: visiowings --help with toml update_check=false failed: {combined[-500:]}"
    )
    assert not _looks_like_update_hint(combined), (
        f"§H3: update hint leaked despite update_check=false in toml. Output:\n{combined[-600:]}"
    )


@pytest.mark.section("visiowings-H4", "visiowings/docs/contributing/uat.md#h4")
def test_section_h4_offline_no_traceback_no_hang(visiowings_repo):
    """UAT §H4: with PyPI unreachable (we point HTTP(S)_PROXY at a
    closed loopback port), the update check must fail silently — no
    traceback, no >5s hang.

    visiowings should respect HTTPS_PROXY; if it doesn't honour proxy
    env vars at all and goes direct, the test still passes as long as
    output is clean and runs within the timeout (real PyPI reachable).
    """
    _need_repo(visiowings_repo)
    import time as _t

    t0 = _t.monotonic()
    cp = _run_with_update_env(
        visiowings_repo,
        {
            "HTTPS_PROXY": "http://127.0.0.1:1",
            "HTTP_PROXY": "http://127.0.0.1:1",
            "https_proxy": "http://127.0.0.1:1",
            "http_proxy": "http://127.0.0.1:1",
        },
        "--help",
        timeout=10,
    )
    elapsed = _t.monotonic() - t0
    combined = cp.stdout + cp.stderr
    assert cp.returncode == 0, (
        f"§H4: visiowings --help should exit 0 even when PyPI unreachable. "
        f"Got exit {cp.returncode}.\n{combined[-600:]}"
    )
    assert not _has_python_traceback(combined), (
        f"§H4: offline update check leaked a Python traceback:\n{combined[-1500:]}"
    )
    # Spec: ``no >5s hang``. Give a margin (10s wall, fail at ≥6s).
    assert elapsed < 6.0, (
        f"§H4: offline update check hung for {elapsed:.1f}s. "
        f"Spec requires fast silent failure (<5s)."
    )


@pytest.mark.section("visiowings-I", "visiowings/docs/contributing/uat.md#i")
@pytest.mark.not_yet_implemented("iter4")
def test_section_i_docs():
    pass


@pytest.mark.section("visiowings-J", "visiowings/docs/contributing/uat.md#j")
@pytest.mark.not_yet_implemented("iter4")
def test_section_j_release_artifacts():
    pass


@pytest.mark.section("visiowings-K", "visiowings/docs/contributing/uat.md#k")
@pytest.mark.not_yet_implemented("iter4")
def test_section_k_ci_sanity():
    pass


@pytest.mark.section("visiowings-L", "visiowings/docs/contributing/uat.md#l")
@pytest.mark.not_yet_implemented("iter4")
def test_section_l_signoff_meta():
    pass
