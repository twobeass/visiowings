"""Tests for visiowings._logging.setup_logging."""

from __future__ import annotations

import io
import logging

import pytest

from visiowings._logging import ColoringFormatter, setup_logging


@pytest.fixture(autouse=True)
def reset_root_logger():
    """Restore root logger state after every test."""

    root = logging.getLogger()
    handlers = list(root.handlers)
    level = root.level
    yield
    for h in list(root.handlers):
        root.removeHandler(h)
    for h in handlers:
        root.addHandler(h)
    root.setLevel(level)


def test_setup_installs_single_handler():
    stream = io.StringIO()
    setup_logging(debug=False, stream=stream)
    root = logging.getLogger()
    assert len(root.handlers) == 1


def test_setup_idempotent_does_not_accumulate_handlers():
    stream = io.StringIO()
    setup_logging(debug=False, stream=stream)
    setup_logging(debug=True, stream=stream)
    setup_logging(debug=False, stream=stream)
    assert len(logging.getLogger().handlers) == 1


def test_debug_flag_enables_debug_level():
    stream = io.StringIO()
    setup_logging(debug=True, stream=stream)
    assert logging.getLogger().level == logging.DEBUG


def test_default_uses_info_level():
    stream = io.StringIO()
    setup_logging(debug=False, stream=stream)
    assert logging.getLogger().level == logging.INFO


def test_messages_are_written_to_stream():
    stream = io.StringIO()
    setup_logging(debug=False, stream=stream)
    logging.getLogger("test").info("hello world")
    assert "hello world" in stream.getvalue()


def test_debug_format_includes_logger_name():
    stream = io.StringIO()
    setup_logging(debug=True, stream=stream)
    logging.getLogger("visiowings.foo").debug("hi")
    output = stream.getvalue()
    assert "DEBUG" in output
    assert "visiowings.foo" in output


def test_no_color_env_disables_color(monkeypatch):
    monkeypatch.setenv("NO_COLOR", "1")
    stream = io.StringIO()
    stream.isatty = lambda: True  # type: ignore[method-assign]
    setup_logging(debug=False, stream=stream)
    logging.getLogger().error("oops")
    output = stream.getvalue()
    assert "\x1b[" not in output  # no ANSI escapes


def test_force_color_env_enables_color(monkeypatch):
    monkeypatch.delenv("NO_COLOR", raising=False)
    monkeypatch.setenv("FORCE_COLOR", "1")
    stream = io.StringIO()
    stream.isatty = lambda: False  # type: ignore[method-assign]
    setup_logging(debug=False, stream=stream)
    logging.getLogger().error("oops")
    output = stream.getvalue()
    assert "\x1b[31m" in output  # red for ERROR


def test_coloring_formatter_passthrough_when_disabled():
    fmt = ColoringFormatter(fmt="%(message)s", color=False)
    record = logging.LogRecord(
        name="x", level=logging.ERROR, pathname="x", lineno=1,
        msg="boom", args=(), exc_info=None,
    )
    assert fmt.format(record) == "boom"


def test_coloring_formatter_no_style_for_unknown_level():
    fmt = ColoringFormatter(fmt="%(message)s", color=True)
    record = logging.LogRecord(
        name="x", level=logging.NOTSET, pathname="x", lineno=1,
        msg="meh", args=(), exc_info=None,
    )
    # NOTSET (0) has no style entry, output should be unchanged.
    assert fmt.format(record) == "meh"
