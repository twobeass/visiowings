"""Tests for CLI validation helpers."""

from __future__ import annotations

from pathlib import Path

import pytest

from visiowings.cli import (
    _build_parser,
    _validate_codepage,
    _validate_readable_dir,
    _validate_visio_file,
    _validate_writable_dir,
)
from visiowings.exceptions import (
    DocumentNotFoundError,
    InvalidVisioFileError,
    UnsupportedEncodingError,
    VisiowingsError,
)


# --------------------------------------------------------------------------- #
# _validate_visio_file
# --------------------------------------------------------------------------- #
class TestValidateVisioFile:
    def test_rejects_unknown_suffix(self, tmp_path):
        bad = tmp_path / "foo.txt"
        bad.write_text("nope")
        with pytest.raises(InvalidVisioFileError) as info:
            _validate_visio_file(bad)
        assert ".vsdm" in info.value.message

    def test_rejects_missing_file(self, tmp_path):
        with pytest.raises(DocumentNotFoundError):
            _validate_visio_file(tmp_path / "missing.vsdm")

    def test_accepts_existing_vsdm(self, tmp_path):
        good = tmp_path / "ok.vsdm"
        good.write_bytes(b"\x00")  # contents irrelevant
        result = _validate_visio_file(good)
        assert result == good.resolve()

    @pytest.mark.parametrize(
        "suffix",
        [".vsd", ".vsdx", ".vsdm", ".vstx", ".vstm", ".vssm", ".vssx"],
    )
    def test_all_supported_suffixes_accepted(self, tmp_path, suffix):
        path = tmp_path / f"file{suffix}"
        path.write_bytes(b"\x00")
        assert _validate_visio_file(path).suffix == suffix


# --------------------------------------------------------------------------- #
# _validate_codepage
# --------------------------------------------------------------------------- #
class TestValidateCodepage:
    def test_none_passes_through(self):
        assert _validate_codepage(None) is None

    def test_empty_string_passes_through(self):
        assert _validate_codepage("") is None

    @pytest.mark.parametrize("name", ["cp1252", "cp1251", "cp932", "utf-8", "ascii"])
    def test_known_codepage_returned_unchanged(self, name):
        assert _validate_codepage(name) == name

    def test_unknown_codepage_raises(self):
        with pytest.raises(UnsupportedEncodingError):
            _validate_codepage("not-a-codepage")


# --------------------------------------------------------------------------- #
# _validate_writable_dir / _validate_readable_dir
# --------------------------------------------------------------------------- #
class TestDirectoryValidation:
    def test_creates_missing_writable_dir(self, tmp_path):
        target = tmp_path / "new"
        result = _validate_writable_dir(target, label="--output")
        assert result.is_dir()

    def test_writable_dir_returns_resolved_path(self, tmp_path):
        target = tmp_path / "out"
        target.mkdir()
        result = _validate_writable_dir(target, label="--output")
        assert result == target.resolve()

    def test_writable_dir_rejects_when_touch_raises(self, tmp_path, monkeypatch):
        """Iter4 §G4: Windows ACL deny rules are invisible to `os.access`
        but DO raise PermissionError on a real write. The probe must catch
        that and raise `VisiowingsError`, so a per-batch export never
        starts against an unwritable output dir."""

        target = tmp_path / "ro"
        target.mkdir()

        original_touch = Path.touch

        def _refusing_touch(self_path, *a, **kw):
            if self_path.name == ".visiowings-write-probe":
                raise PermissionError("simulated ACL deny")
            return original_touch(self_path, *a, **kw)

        monkeypatch.setattr(Path, "touch", _refusing_touch)

        with pytest.raises(VisiowingsError) as info:
            _validate_writable_dir(target, label="--output")
        assert "not writable" in str(info.value)
        # No leftover probe file.
        assert not (target / ".visiowings-write-probe").exists()

    def test_readable_dir_rejects_missing(self, tmp_path):
        with pytest.raises(VisiowingsError):
            _validate_readable_dir(tmp_path / "missing", label="--input")

    def test_readable_dir_rejects_file(self, tmp_path):
        f = tmp_path / "file.txt"
        f.write_text("x")
        with pytest.raises(VisiowingsError):
            _validate_readable_dir(f, label="--input")

    def test_readable_dir_returns_resolved(self, tmp_path):
        d = tmp_path / "in"
        d.mkdir()
        assert _validate_readable_dir(d, label="--input") == d.resolve()


# --------------------------------------------------------------------------- #
# CLI top-level behaviour
# --------------------------------------------------------------------------- #
class TestParser:
    def test_version_flag_prints_version(self, capsys):
        parser = _build_parser()
        with pytest.raises(SystemExit) as info:
            parser.parse_args(["--version"])
        assert info.value.code == 0
        captured = capsys.readouterr()
        assert "visiowings" in captured.out

    def test_help_does_not_explode(self, capsys):
        parser = _build_parser()
        with pytest.raises(SystemExit) as info:
            parser.parse_args(["--help"])
        assert info.value.code == 0

    def test_edit_does_not_require_file_on_parser_level(self):
        # --file moved out of `required=True` so a `.visiowings.toml` config
        # can supply the value. The CLI's main() raises later when neither
        # flag nor config provides it (covered separately).
        parser = _build_parser()
        ns = parser.parse_args(["edit"])
        assert ns.file is None


def test_main_returns_nonzero_on_validation_error(capsys, tmp_path, monkeypatch):
    """End-to-end: bad --file triggers VisiowingsError catch → exit 1."""

    from visiowings import cli

    bad = tmp_path / "foo.txt"
    bad.write_text("x")
    rc = cli.main(["edit", "--file", str(bad)])
    assert rc == 1
    captured = capsys.readouterr()
    assert "❌" in captured.err


# --------------------------------------------------------------------------- #
# UAT §C1 — cp1252-safe streams
# --------------------------------------------------------------------------- #
class TestForceUtf8Streams:
    """`_force_utf8_streams` must make CLI output safe on legacy ANSI consoles.

    The real-world bug: Windows German/French defaults give stdout a cp1252
    codec which cannot encode the emoji we use in banners. `cmd_init`
    crashed before writing `.visiowings.toml`.
    """

    def test_calls_reconfigure_with_utf8_replace(self, monkeypatch):
        from visiowings import cli

        calls: list[dict] = []

        class _FakeStream:
            def reconfigure(self, **kwargs):
                calls.append(kwargs)

        monkeypatch.setattr(cli.sys, "stdout", _FakeStream())
        monkeypatch.setattr(cli.sys, "stderr", _FakeStream())

        cli._force_utf8_streams()

        assert len(calls) == 2  # one per stream
        for c in calls:
            assert c == {"encoding": "utf-8", "errors": "replace"}

    def test_silently_skips_streams_without_reconfigure(self, monkeypatch):
        """Captured streams (pytest's capsys) lack `.reconfigure`; that's fine."""

        from visiowings import cli

        class _DumbStream:
            pass  # no reconfigure attribute

        monkeypatch.setattr(cli.sys, "stdout", _DumbStream())
        monkeypatch.setattr(cli.sys, "stderr", _DumbStream())

        # Must not raise.
        cli._force_utf8_streams()

    def test_swallows_reconfigure_errors(self, monkeypatch):
        from visiowings import cli

        class _AngryStream:
            def reconfigure(self, **kwargs):
                raise OSError("detached stream")

        monkeypatch.setattr(cli.sys, "stdout", _AngryStream())
        monkeypatch.setattr(cli.sys, "stderr", _AngryStream())

        # Must not raise.
        cli._force_utf8_streams()


# --------------------------------------------------------------------------- #
# UAT §G4 — export must exit non-zero when per-document writes fail
# --------------------------------------------------------------------------- #
class TestExportExitCodeOnFailure:
    """`cmd_export` must translate the exporter's per-doc failure log
    into a `VisiowingsError`, so `main()` returns exit 1 even when only
    *some* of the documents in the batch failed (or all of them did
    without raising at the batch level)."""

    def test_main_returns_one_when_exporter_records_failure(self, monkeypatch, tmp_path, capsys):
        """End-to-end: a stub exporter that records a failure → main → 1."""

        from visiowings import cli

        vsdm = tmp_path / "doc.vsdm"
        vsdm.write_bytes(b"")  # extension is what we validate

        class _StubExporter:
            def __init__(self, *_a, **_kw):
                # Mirror the public attribute the CLI reads.
                self.last_export_failures = []

            def connect_to_visio(self):
                return True

            def export_modules(self, _output_dir, last_hashes=None):
                self.last_export_failures = [
                    {"document": "Drawing1", "error": "PermissionError: [WinError 5]"}
                ]
                return {}, {}

        # Patch the lazy import target.
        import sys as _sys

        monkeypatch.setattr(
            _sys.modules["visiowings.vba_export"], "VisioVBAExporter", _StubExporter
        )

        rc = cli.main(["export", "--file", str(vsdm), "--output", str(tmp_path)])
        assert rc == 1
        out = capsys.readouterr()
        assert "❌" in out.err
        assert "Export failed" in out.err
        assert "Drawing1" in out.err
        assert "PermissionError" in out.err

    def test_main_returns_zero_when_no_failures_recorded(self, monkeypatch, tmp_path, capsys):
        """Sanity guard: when the stub does NOT record a failure, exit 0."""

        from visiowings import cli

        vsdm = tmp_path / "doc.vsdm"
        vsdm.write_bytes(b"")

        class _StubExporter:
            def __init__(self, *_a, **_kw):
                self.last_export_failures = []

            def connect_to_visio(self):
                return True

            def export_modules(self, _output_dir, last_hashes=None):
                return {"Drawing1": ["Module1.bas"]}, {"Drawing1": "hash"}

        import sys as _sys

        monkeypatch.setattr(
            _sys.modules["visiowings.vba_export"], "VisioVBAExporter", _StubExporter
        )

        rc = cli.main(["export", "--file", str(vsdm), "--output", str(tmp_path)])
        assert rc == 0
