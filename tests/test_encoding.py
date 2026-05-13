"""Tests for visiowings.encoding (LCID resolution + BOM-aware decoding)."""

from __future__ import annotations

import codecs
from unittest.mock import MagicMock

import pytest

from visiowings.encoding import (
    DEFAULT_CODEPAGE,
    LCID_TO_CODEPAGE,
    get_encoding_from_document,
    resolve_encoding,
)


# --------------------------------------------------------------------------- #
# LCID -> codepage mapping
# --------------------------------------------------------------------------- #
class TestLcidMapping:
    def test_default_codepage_is_resolvable(self):
        assert codecs.lookup(DEFAULT_CODEPAGE)

    @pytest.mark.parametrize("lcid,expected", sorted(LCID_TO_CODEPAGE.items()))
    def test_every_mapped_codepage_is_known_to_python(self, lcid, expected):
        # Every codepage in the table must be a valid Python codec name.
        info = codecs.lookup(expected)
        assert info.name  # smoke check

    def test_us_english_lcid_maps_to_cp1252(self):
        assert LCID_TO_CODEPAGE[1033] == "cp1252"

    def test_russian_lcid_maps_to_cp1251(self):
        assert LCID_TO_CODEPAGE[1049] == "cp1251"

    def test_japanese_lcid_maps_to_cp932(self):
        assert LCID_TO_CODEPAGE[1041] == "cp932"

    def test_simplified_chinese_lcid_maps_to_cp936(self):
        assert LCID_TO_CODEPAGE[2052] == "cp936"

    def test_korean_lcid_maps_to_cp949(self):
        assert LCID_TO_CODEPAGE[1042] == "cp949"

    def test_thai_lcid_maps_to_cp874(self):
        assert LCID_TO_CODEPAGE[1054] == "cp874"


# --------------------------------------------------------------------------- #
# get_encoding_from_document
# --------------------------------------------------------------------------- #
class TestGetEncodingFromDocument:
    def test_returns_mapped_codepage_for_known_lcid(self):
        doc = MagicMock()
        doc.Language = 1031  # de-DE
        assert get_encoding_from_document(doc) == "cp1252"

    def test_returns_none_for_unknown_lcid(self):
        doc = MagicMock()
        doc.Language = 99999  # not in table
        assert get_encoding_from_document(doc) is None

    def test_returns_none_when_language_attr_raises(self):
        doc = MagicMock()
        type(doc).Language = property(lambda self: (_ for _ in ()).throw(AttributeError("nope")))
        assert get_encoding_from_document(doc) is None

    def test_zero_lcid_returns_none(self):
        doc = MagicMock()
        doc.Language = 0
        assert get_encoding_from_document(doc) is None


# --------------------------------------------------------------------------- #
# resolve_encoding (priority order)
# --------------------------------------------------------------------------- #
class TestResolveEncoding:
    def test_user_codepage_wins_over_document(self):
        doc = MagicMock()
        doc.Language = 1031  # would map to cp1252
        assert resolve_encoding(doc, user_codepage="cp1251") == "cp1251"

    def test_document_language_wins_over_default(self):
        doc = MagicMock()
        doc.Language = 1049  # ru-RU -> cp1251
        assert resolve_encoding(doc) == "cp1251"

    def test_falls_back_to_default_when_document_unmapped(self):
        doc = MagicMock()
        doc.Language = 99999
        assert resolve_encoding(doc) == DEFAULT_CODEPAGE

    def test_falls_back_to_default_when_no_document(self):
        assert resolve_encoding(document=None) == DEFAULT_CODEPAGE

    def test_user_codepage_used_even_when_document_is_none(self):
        assert resolve_encoding(document=None, user_codepage="cp932") == "cp932"


# --------------------------------------------------------------------------- #
# BOM detection (used by VBA importer to clean files before decoding)
# --------------------------------------------------------------------------- #
@pytest.mark.parametrize(
    "bom,encoding,sample",
    [
        (codecs.BOM_UTF8, "utf-8-sig", "Hello"),
        (codecs.BOM_UTF16_LE, "utf-16", "Hello"),
        (codecs.BOM_UTF16_BE, "utf-16", "Hello"),
    ],
)
def test_python_can_round_trip_bom_marked_text(bom, encoding, sample, tmp_path):
    """Sanity check that BOM-marked content decodes via standard codecs.

    Phase D will add an explicit BOM-stripping helper, but this test
    locks in the underlying codec behaviour we will rely on.
    """

    path = tmp_path / "fixture.bas"
    if encoding == "utf-16":
        # codecs.BOM_UTF16_* are raw bytes; pair them with utf-16-le/be encodings.
        suffix = "le" if bom == codecs.BOM_UTF16_LE else "be"
        encoded = bom + sample.encode(f"utf-16-{suffix}")
        path.write_bytes(encoded)
    else:
        path.write_bytes(bom + sample.encode("utf-8"))

    decoded = path.read_text(encoding=encoding)
    assert decoded.replace("﻿", "") == sample
