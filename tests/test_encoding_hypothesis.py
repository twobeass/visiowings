"""Property-based encoding tests with Hypothesis."""

from __future__ import annotations

import pytest
from hypothesis import HealthCheck, given, settings
from hypothesis import strategies as st

from visiowings.encoding import LCID_TO_CODEPAGE

UNIQUE_CODEPAGES = sorted(set(LCID_TO_CODEPAGE.values()))


@settings(
    max_examples=50,
    deadline=None,
    suppress_health_check=[HealthCheck.too_slow, HealthCheck.filter_too_much],
)
@given(text=st.text(max_size=200))
@pytest.mark.parametrize("codepage", UNIQUE_CODEPAGES)
def test_encode_with_replace_is_decodable(codepage, text):
    """For any text, encode(errors='replace') -> decode() must not raise."""

    encoded = text.encode(codepage, errors="replace")
    decoded = encoded.decode(codepage)
    assert isinstance(decoded, str)


@settings(max_examples=200, deadline=None)
@given(
    # ASCII subset: everything in this range is in every cp* codepage.
    text=st.text(alphabet=st.characters(min_codepoint=0x20, max_codepoint=0x7E), max_size=200),
)
@pytest.mark.parametrize("codepage", UNIQUE_CODEPAGES)
def test_ascii_round_trip_is_lossless_in_every_codepage(codepage, text):
    """Pure ASCII must round-trip through every codepage we declare."""

    encoded = text.encode(codepage)
    decoded = encoded.decode(codepage)
    assert decoded == text


@settings(
    max_examples=50,
    deadline=None,
    suppress_health_check=[HealthCheck.filter_too_much],
)
@given(text=st.text(max_size=200))
def test_utf8_round_trip_is_lossless(text):
    assert text.encode("utf-8").decode("utf-8") == text
