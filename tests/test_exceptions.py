"""Tests for visiowings.exceptions."""

from __future__ import annotations

import pytest

from visiowings.exceptions import (
    COMConnectionError,
    DocumentNotFoundError,
    InvalidVisioFileError,
    UnsupportedEncodingError,
    VBAImportError,
    VisioNotRunningError,
    VisiowingsError,
)


def test_visiowings_error_is_exception_subclass():
    assert issubclass(VisiowingsError, Exception)


def test_message_attribute_falls_back_to_docstring():
    err = VisioNotRunningError()
    assert err.message  # non-empty
    assert "Visio" in err.message


def test_document_not_found_lists_open_documents():
    err = DocumentNotFoundError("missing.vsdm", available=["a.vsdm", "b.vsdx"])
    assert err.requested == "missing.vsdm"
    assert "a.vsdm" in err.message
    assert "b.vsdx" in err.message


def test_document_not_found_handles_empty_list():
    err = DocumentNotFoundError("missing.vsdm", available=[])
    assert "No documents are currently open" in err.message


def test_unsupported_encoding_includes_suggestions():
    err = UnsupportedEncodingError("xyz")
    assert "xyz" in err.message
    assert "cp1252" in err.message  # at least one suggestion


def test_com_connection_error_carries_attempt_count_and_cause():
    cause = RuntimeError("boom")
    err = COMConnectionError(3, last_error=cause)
    assert err.attempts == 3
    assert err.last_error is cause
    assert "3" in err.message


def test_vba_import_error_includes_file_and_reason():
    err = VBAImportError("Module1.bas", "syntax error on line 10")
    assert "Module1.bas" in err.message
    assert "syntax error on line 10" in err.message


def test_invalid_visio_file_lists_supported_extensions():
    err = InvalidVisioFileError("foo.txt")
    assert "foo.txt" in err.message
    assert ".vsdm" in err.message
    assert ".vsdx" in err.message


def test_invalid_visio_file_extension_set_is_complete():
    # The CLI relies on this list to validate input. Lock it in to prevent
    # silent shrinking.
    expected = {".vsd", ".vsdx", ".vsdm", ".vstx", ".vstm", ".vssm", ".vssx"}
    assert set(InvalidVisioFileError.SUPPORTED_SUFFIXES) == expected


def test_can_be_raised_and_caught():
    with pytest.raises(VisiowingsError):
        raise VisioNotRunningError()
