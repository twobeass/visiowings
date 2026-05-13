"""Tests for visiowings._retry.retry_with_backoff."""

from __future__ import annotations

from unittest.mock import patch

import pytest

from visiowings._retry import retry_with_backoff
from visiowings.exceptions import COMConnectionError


@pytest.fixture(autouse=True)
def patch_sleep():
    with patch("visiowings._retry.time.sleep") as sleep:
        yield sleep


def test_returns_immediately_on_success(patch_sleep):
    @retry_with_backoff(max_attempts=3)
    def ok():
        return 42

    assert ok() == 42
    patch_sleep.assert_not_called()


def test_retries_until_success(patch_sleep):
    calls = {"n": 0}

    @retry_with_backoff(max_attempts=3, exceptions=(RuntimeError,))
    def flaky():
        calls["n"] += 1
        if calls["n"] < 3:
            raise RuntimeError("nope")
        return "ok"

    assert flaky() == "ok"
    assert calls["n"] == 3
    assert patch_sleep.call_count == 2  # one sleep between each attempt


def test_raises_com_connection_error_after_exhaustion(patch_sleep):
    @retry_with_backoff(max_attempts=3, exceptions=(RuntimeError,))
    def always_fails():
        raise RuntimeError("boom")

    with pytest.raises(COMConnectionError) as info:
        always_fails()
    assert info.value.attempts == 3
    assert isinstance(info.value.__cause__, RuntimeError)


def test_does_not_retry_unmatched_exception(patch_sleep):
    @retry_with_backoff(max_attempts=5, exceptions=(RuntimeError,))
    def wrong_kind():
        raise ValueError("specific")

    with pytest.raises(ValueError):
        wrong_kind()
    patch_sleep.assert_not_called()


def test_exponential_backoff_delays(patch_sleep):
    @retry_with_backoff(
        max_attempts=4,
        initial_delay=0.5,
        multiplier=2.0,
        max_delay=10.0,
        exceptions=(RuntimeError,),
    )
    def always_fails():
        raise RuntimeError("boom")

    with pytest.raises(COMConnectionError):
        always_fails()
    # 4 attempts -> 3 sleeps with delays 0.5, 1.0, 2.0
    delays = [call.args[0] for call in patch_sleep.call_args_list]
    assert delays == [0.5, 1.0, 2.0]


def test_max_delay_caps_backoff(patch_sleep):
    @retry_with_backoff(
        max_attempts=5,
        initial_delay=2.0,
        multiplier=10.0,
        max_delay=4.0,
        exceptions=(RuntimeError,),
    )
    def always_fails():
        raise RuntimeError("boom")

    with pytest.raises(COMConnectionError):
        always_fails()
    delays = [call.args[0] for call in patch_sleep.call_args_list]
    # 2.0 -> 4.0 (capped) -> 4.0 -> 4.0
    assert delays == [2.0, 4.0, 4.0, 4.0]


def test_invalid_max_attempts_raises():
    with pytest.raises(ValueError):
        retry_with_backoff(max_attempts=0)


def test_custom_raise_on_failure(patch_sleep):
    @retry_with_backoff(
        max_attempts=2,
        exceptions=(RuntimeError,),
        raise_on_failure=ValueError,
    )
    def always_fails():
        raise RuntimeError("boom")

    with pytest.raises(ValueError) as info:
        always_fails()
    assert isinstance(info.value.__cause__, RuntimeError)
