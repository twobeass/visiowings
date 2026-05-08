"""BOM-aware decoding in VisioVBAImporter._decode_with_bom_detection."""

from __future__ import annotations

import codecs

import pytest

from visiowings.vba_import import VisioVBAImporter


def _write(tmp_path, name: str, raw: bytes):
    path = tmp_path / name
    path.write_bytes(raw)
    return path


def test_utf8_bom_is_stripped(tmp_path):
    raw = codecs.BOM_UTF8 + b"Sub Foo()\nEnd Sub\n"
    path = _write(tmp_path, "utf8bom.bas", raw)
    text = VisioVBAImporter._decode_with_bom_detection(path, "cp1252")
    assert text.startswith("Sub Foo()")
    assert "﻿" not in text


def test_utf16_le_bom_is_stripped(tmp_path):
    raw = codecs.BOM_UTF16_LE + "Sub Foo()\nEnd Sub\n".encode("utf-16-le")
    path = _write(tmp_path, "utf16le.bas", raw)
    text = VisioVBAImporter._decode_with_bom_detection(path, "cp1252")
    assert text.startswith("Sub Foo()")
    assert "﻿" not in text


def test_utf16_be_bom_is_stripped(tmp_path):
    raw = codecs.BOM_UTF16_BE + "Sub Foo()\nEnd Sub\n".encode("utf-16-be")
    path = _write(tmp_path, "utf16be.bas", raw)
    text = VisioVBAImporter._decode_with_bom_detection(path, "cp1252")
    assert text.startswith("Sub Foo()")


def test_no_bom_utf8_decoded_as_utf8(tmp_path):
    raw = "Sub Foo()\n' café\nEnd Sub\n".encode("utf-8")
    path = _write(tmp_path, "nobom.bas", raw)
    text = VisioVBAImporter._decode_with_bom_detection(path, "cp1252")
    assert "café" in text


def test_falls_back_to_codepage_when_utf8_invalid(tmp_path):
    # cp1251-encoded Cyrillic bytes that are not valid UTF-8.
    raw = "Привет".encode("cp1251")
    path = _write(tmp_path, "cyrillic.bas", raw)
    text = VisioVBAImporter._decode_with_bom_detection(path, "cp1251")
    assert text == "Привет"


def test_replacement_char_when_neither_codec_works(tmp_path):
    """Bytes that are not valid UTF-8 and not a BOM still decode without raising.

    We pick a leading byte (0x80) that:
      - Is not the start of any BOM (UTF-8/16-LE/16-BE).
      - Is invalid as the first byte of a UTF-8 sequence (continuation byte).
      - Is mappable in cp1252 (so the fallback branch produces text).
    """

    raw = b"\x80hello"
    path = _write(tmp_path, "garbage.bas", raw)
    text = VisioVBAImporter._decode_with_bom_detection(path, "cp1252")
    assert isinstance(text, str)
    assert "hello" in text


@pytest.mark.parametrize(
    "bom",
    [codecs.BOM_UTF8, codecs.BOM_UTF16_LE, codecs.BOM_UTF16_BE],
)
def test_empty_file_with_only_bom_returns_empty(tmp_path, bom):
    path = _write(tmp_path, "only-bom.bas", bom)
    text = VisioVBAImporter._decode_with_bom_detection(path, "cp1252")
    assert text == ""
