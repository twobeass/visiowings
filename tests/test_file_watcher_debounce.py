"""Threading-safe debounce in VBAFileHandler._record_change."""

from __future__ import annotations

import threading
from unittest.mock import MagicMock

from visiowings.file_watcher import VBAFileHandler


def _handler():
    return VBAFileHandler(MagicMock(), MagicMock())


def test_first_event_is_recorded():
    h = _handler()
    assert h._record_change("/a.bas", now=10.0) is True


def test_repeat_event_within_window_is_debounced():
    h = _handler()
    assert h._record_change("/a.bas", now=10.0) is True
    assert h._record_change("/a.bas", now=10.5) is False


def test_event_after_window_is_recorded():
    h = _handler()
    assert h._record_change("/a.bas", now=10.0) is True
    assert h._record_change("/a.bas", now=11.5) is True


def test_different_files_not_cross_debounced():
    h = _handler()
    assert h._record_change("/a.bas", now=10.0) is True
    assert h._record_change("/b.bas", now=10.0) is True


def test_lru_cap_evicts_old_entries():
    # Force the cap down to keep the test fast.
    from visiowings import file_watcher as fw

    h = _handler()
    cap = 5
    real_cap = fw._DEBOUNCE_MAX_ENTRIES
    fw._DEBOUNCE_MAX_ENTRIES = cap
    try:
        # Spread events 100s apart so debounce never kicks in. Each path is
        # unique, so all events should be recorded.
        for i in range(cap + 3):
            assert h._record_change(f"/file_{i}.bas", now=100.0 + i * 100) is True
        assert len(h._last_modified) == cap
        # The oldest 3 entries were evicted by the LRU cap.
        for i in range(3):
            assert f"/file_{i}.bas" not in h._last_modified
    finally:
        fw._DEBOUNCE_MAX_ENTRIES = real_cap


def test_concurrent_events_do_not_corrupt_dict():
    """Hammer _record_change from many threads. No exception, no lost rows."""

    h = _handler()
    files = [f"/concurrent_{i}.bas" for i in range(50)]
    barrier = threading.Barrier(len(files))

    def worker(path: str) -> None:
        barrier.wait()
        for offset in range(20):
            h._record_change(path, now=offset * 2.0)  # always >1s apart

    threads = [threading.Thread(target=worker, args=(p,)) for p in files]
    for t in threads:
        t.start()
    for t in threads:
        t.join()

    assert set(h._last_modified) == set(files)
