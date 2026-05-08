"""Tests for the opt-out PyPI update check."""

from __future__ import annotations

import json
from pathlib import Path
from unittest.mock import patch

import pytest

from visiowings import _update_check


@pytest.fixture(autouse=True)
def isolate_cache(tmp_path, monkeypatch):
    """Redirect the cache file to tmp_path so tests don't touch ~/.cache."""

    cache_file = tmp_path / "update_check.json"
    monkeypatch.setattr(_update_check, "_CACHE_DIR", tmp_path)
    monkeypatch.setattr(_update_check, "_CACHE_FILE", cache_file)
    monkeypatch.delenv("VISIOWINGS_NO_UPDATE_CHECK", raising=False)
    monkeypatch.delenv("CI", raising=False)
    yield


def test_disabled_via_env_var(monkeypatch):
    monkeypatch.setenv("VISIOWINGS_NO_UPDATE_CHECK", "1")
    assert _update_check._is_disabled() is True


def test_disabled_in_ci(monkeypatch):
    monkeypatch.setenv("CI", "true")
    assert _update_check._is_disabled() is True


def test_not_disabled_by_default():
    assert _update_check._is_disabled() is False


@pytest.mark.parametrize(
    "current,latest,expected",
    [
        ("0.6.1", "0.7.0", True),
        ("0.6.1", "1.0.0", True),
        ("0.6.1", "0.6.1", False),
        ("1.0.0", "0.9.9", False),
        ("0.6.1", "0.6.2", True),
    ],
)
def test_is_outdated(current, latest, expected):
    assert _update_check._is_outdated(current, latest) is expected


def test_parse_version_handles_pre_release_tags():
    # 0.7.0rc1 stops at the rc-tagged segment, so it parses as (0, 7).
    assert _update_check._parse_version("0.7.0rc1") == (0, 7)
    # Hyphen / underscore are normalised to dots before parsing.
    assert _update_check._parse_version("1.0.0-beta") == (1, 0, 0)


def test_pre_release_is_older_than_release():
    """Real-world expectation: 0.7.0rc1 should be considered older than 0.7.0."""

    assert _update_check._is_outdated("0.7.0rc1", "0.7.0") is True


def test_load_cache_returns_none_when_missing():
    assert _update_check._load_cache() is None


def test_save_and_load_cache_round_trip():
    payload = {"latest": "0.9.0", "checked_at": 1234567890}
    _update_check._save_cache(payload)
    loaded = _update_check._load_cache()
    assert loaded == payload


def test_load_cache_skips_stale_entries(monkeypatch):
    payload = {"latest": "0.9.0", "checked_at": 0}
    _update_check._save_cache(payload)
    # Pretend the cache file is from 1990.
    import os, time
    mtime = time.time() - (_update_check._CACHE_TTL_SECONDS + 1)
    os.utime(_update_check._CACHE_FILE, (mtime, mtime))
    assert _update_check._load_cache() is None


def test_check_and_notify_writes_hint_when_outdated(monkeypatch, capsys):
    monkeypatch.setattr(_update_check, "_fetch_latest_version", lambda: "9.9.9")
    monkeypatch.setattr(_update_check, "__version__", "0.6.1")
    _update_check._check_and_notify()
    captured = capsys.readouterr()
    assert "9.9.9" in captured.err
    assert "pipx upgrade" in captured.err or "pip install -U" in captured.err


def test_check_and_notify_silent_when_current(monkeypatch, capsys):
    monkeypatch.setattr(_update_check, "_fetch_latest_version", lambda: "0.6.1")
    monkeypatch.setattr(_update_check, "__version__", "0.6.1")
    _update_check._check_and_notify()
    captured = capsys.readouterr()
    assert captured.err == ""


def test_schedule_async_skips_when_disabled(monkeypatch):
    monkeypatch.setenv("VISIOWINGS_NO_UPDATE_CHECK", "1")
    with patch.object(_update_check, "_check_and_notify") as check:
        _update_check.schedule_async()
    check.assert_not_called()


def test_schedule_async_runs_in_background():
    # When enabled the function spawns a daemon thread and returns.
    _update_check.schedule_async()  # should not raise
