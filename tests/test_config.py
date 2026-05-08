"""Tests for visiowings.config — .visiowings.toml load/find/write."""

from __future__ import annotations

from pathlib import Path

import pytest

from visiowings.config import (
    CONFIG_FILENAME,
    VisiowingsConfig,
    find_config,
    load_config,
    write_config,
)


def test_default_config_is_empty():
    cfg = VisiowingsConfig()
    assert cfg.file is None
    assert cfg.bidirectional is False
    assert cfg.codepage is None


def test_to_toml_includes_keys():
    cfg = VisiowingsConfig(
        file="main.vsdm",
        output="vba",
        codepage="cp1252",
        bidirectional=True,
        rubberduck=True,
    )
    rendered = cfg.to_toml()
    assert 'file = "main.vsdm"' in rendered
    assert 'output = "vba"' in rendered
    assert 'codepage = "cp1252"' in rendered
    assert "bidirectional = true" in rendered
    assert "rubberduck = true" in rendered


def test_write_and_load_round_trip(tmp_path):
    cfg = VisiowingsConfig(
        file="drawings/main.vsdm",
        output="vba",
        codepage="cp1251",
        bidirectional=True,
    )
    target = tmp_path / CONFIG_FILENAME
    write_config(cfg, target)
    loaded = load_config(target)
    assert loaded.file == cfg.file
    assert loaded.output == cfg.output
    assert loaded.codepage == cfg.codepage
    assert loaded.bidirectional is True


def test_find_config_walks_up(tmp_path):
    deep = tmp_path / "a" / "b" / "c"
    deep.mkdir(parents=True)
    target = tmp_path / CONFIG_FILENAME
    target.write_text("", encoding="utf-8")
    found = find_config(deep)
    assert found == target


def test_find_config_returns_none_when_absent(tmp_path):
    deep = tmp_path / "x"
    deep.mkdir()
    assert find_config(deep) is None


def test_load_config_auto_discovery_returns_default_when_no_file(tmp_path, monkeypatch):
    """When no path is passed and find_config returns None, we get an empty config."""

    monkeypatch.chdir(tmp_path)
    cfg = load_config()
    assert isinstance(cfg, VisiowingsConfig)
    assert cfg.file is None


def test_load_config_with_explicit_missing_path_raises(tmp_path):
    with pytest.raises(FileNotFoundError):
        load_config(tmp_path / "nope.toml")


def test_unknown_keys_land_in_extras(tmp_path):
    target = tmp_path / CONFIG_FILENAME
    target.write_text("future_flag = true\nfile = \"x.vsdm\"\n", encoding="utf-8")
    cfg = load_config(target)
    assert cfg.file == "x.vsdm"
    assert cfg.extras == {"future_flag": True}
