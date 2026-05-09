"""Tests for CLI validation helpers."""

from __future__ import annotations

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
