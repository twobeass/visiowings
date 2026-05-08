"""Encoding round-trip tests: encode -> bytes -> decode -> identical text.

These tests exercise the codepages declared in ``LCID_TO_CODEPAGE`` to
make sure that VBA modules exported from a Visio document on, say, a
Russian Windows install survive the round-trip through the filesystem.
"""

from __future__ import annotations

import pytest

from visiowings.encoding import LCID_TO_CODEPAGE


# --------------------------------------------------------------------------- #
# Sample strings exercising codepage-specific characters.
# Picking characters that are valid in the given codepage and exotic enough
# to detect mojibake or silent dropping.
# --------------------------------------------------------------------------- #
# Pick characters that exist in each codepage. We deliberately use ASCII
# punctuation (commas, dashes) instead of typographic dashes (U+2014) since
# those are missing from cp932, cp949 and cp1258.
SAMPLES: dict[str, str] = {
    "cp1252": "Hello, cafe - AEOUss - naive resume - ÄÖÜß àéîôü",
    "cp1250": "Prilis zlutoucky kun - asczzo - CSZ - ąśćźż ČŠŽ",
    "cp1251": "Privet, mir! - Moskva - Privetstvuyu - Привет, мир!",
    "cp1253": "Kalimera kosme - Athina - Alphavita - Καλημέρα κόσμε",
    "cp1254": "Turkce - Istanbul - sgucuoc - Türkçe İstanbul",
    "cp1255": "Shalom olam - Yerushalayim - Tel Aviv - שלום עולם",
    "cp1256": "Marhaba bil-alam - al-Qahira - مرحبا بالعالم",
    "cp1257": "Sveika, pasaule - Latviesu valoda - āčēīļš",
    # cp1258 needs decomposed combining diacritics (NFD) for vowels with
    # multiple marks; using NFC precomposed forms (e.g. U+1EBF) would fail.
    "cp1258": "Xin chao, Tieng Viet, dadaeoouw - VN: ô ê",
    "cp874":  "Sawatdi chao lok - phasa thai - สวัสดีชาวโลก",
    "cp932":  "Konnichiwa sekai, nihongo, kanji kana, こんにちは世界",
    "cp936":  "Ni hao shijie, jianti zhongwen, Beijing, 你好世界",
    "cp949":  "Annyeonghaseyo, hangugeo, Seoul, 안녕하세요 한국어",
    "cp950":  "Ni hao shijie, fanti zhongwen, Taipei, 你好世界",
}

# All unique codepages present in the LCID table.
UNIQUE_CODEPAGES = sorted(set(LCID_TO_CODEPAGE.values()))


@pytest.mark.parametrize("codepage", UNIQUE_CODEPAGES)
def test_codepage_is_known_to_python(codepage):
    "".encode(codepage)  # raises LookupError if codepage is unknown


@pytest.mark.parametrize("codepage", UNIQUE_CODEPAGES)
def test_codepage_round_trip_via_tempfile(codepage, tmp_path):
    sample = SAMPLES.get(codepage)
    assert sample is not None, f"No sample defined for {codepage}; please add one."

    path = tmp_path / f"sample-{codepage}.bas"
    path.write_text(sample, encoding=codepage)
    decoded = path.read_text(encoding=codepage)
    assert decoded == sample


@pytest.mark.parametrize(
    "codepage,illegal",
    [
        ("cp1252", "한국어"),  # Korean has no representation in cp1252
        ("cp1251", "日本語"),  # Japanese has no representation in cp1251
        ("cp874", "日本語"),
    ],
)
def test_unrepresentable_chars_replaced_with_errors_replace(codepage, illegal):
    """Errors=replace must produce decodable output without raising."""

    encoded = illegal.encode(codepage, errors="replace")
    decoded = encoded.decode(codepage)
    assert decoded  # non-empty
    assert "?" in decoded  # the replacement char


def test_lcid_table_has_no_orphaned_codepages():
    """Every codepage we map to must actually have a sample, so adding a new
    LCID forces a corresponding round-trip fixture."""

    missing = [cp for cp in UNIQUE_CODEPAGES if cp not in SAMPLES]
    assert not missing, f"Missing round-trip samples for codepages: {missing}"
