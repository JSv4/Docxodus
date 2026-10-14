"""Lifecycle tests — proves session-in-memory persistence is the contract we promise.

These tests load-bear the "very important we are able to keep that session in
memory until explicitly released" requirement: the host's ``SessionRegistry``
holds the parsed ``WordprocessingDocument`` across many Python calls; closing
a session removes it from the registry; closing the host process disposes
everything that's still open.
"""

from __future__ import annotations

import gc

from docx_scalpel import DocxSession, open_session, ping
from docx_scalpel._transport import _Transport, shutdown_host


def _session_count() -> int:
    """Live session count inside the host process."""
    return int(ping().get("sessions", -1))


def test_single_session_survives_many_calls(tour_plan_bytes: bytes) -> None:
    """One session, 30 round-trips, no re-parse cost."""
    session = open_session(tour_plan_bytes)
    try:
        assert session.handle > 0
        assert not session.is_closed
        # Multiple distinct ops should all hit the same in-memory session.
        for _ in range(30):
            proj = session.project()
            assert proj.markdown
            assert len(proj.anchor_index) > 0
    finally:
        session.close()
    assert session.is_closed
    assert _session_count() == 0, "session count should be zero after explicit close"


def test_many_sessions_in_one_host(tour_plan_bytes: bytes) -> None:
    """50 sessions live concurrently in one host process, then all release."""
    sessions: list[DocxSession] = []
    try:
        for _ in range(50):
            sessions.append(open_session(tour_plan_bytes))
        assert _session_count() == 50, (
            f"expected 50 live sessions, host reports {_session_count()}"
        )
        # Each handle is unique.
        handles = {s.handle for s in sessions}
        assert len(handles) == 50, "session handles must be unique"
    finally:
        for s in sessions:
            s.close()
    assert _session_count() == 0, "all sessions should be released after close()"


def test_context_manager_closes_session(tour_plan_bytes: bytes) -> None:
    before = _session_count()
    with open_session(tour_plan_bytes) as session:
        assert _session_count() == before + 1
        assert not session.is_closed
    assert session.is_closed
    assert _session_count() == before, "context-manager exit must close the session"


def test_double_close_is_idempotent(tour_plan_bytes: bytes) -> None:
    session = open_session(tour_plan_bytes)
    session.close()
    session.close()  # must not raise
    assert session.is_closed


def test_calls_after_close_raise(tour_plan_bytes: bytes) -> None:
    import pytest

    session = open_session(tour_plan_bytes)
    session.close()
    with pytest.raises(ValueError, match="closed"):
        session.project()


def test_host_process_is_singleton(tour_plan_bytes: bytes) -> None:
    """Opening multiple sessions reuses the same host PID."""
    s1 = open_session(tour_plan_bytes)
    try:
        pid_before = _Transport._instance._proc.pid  # type: ignore[union-attr]
        s2 = open_session(tour_plan_bytes)
        try:
            pid_after = _Transport._instance._proc.pid  # type: ignore[union-attr]
            assert pid_before == pid_after, (
                "second open_session must reuse the existing host subprocess"
            )
        finally:
            s2.close()
    finally:
        s1.close()


def test_finalizer_falls_back_on_forgotten_session(tour_plan_bytes: bytes) -> None:
    """The __del__ fallback exists so a forgotten session doesn't pin host memory.

    Not the documented path (context manager is), but verifying the safety net works.
    """
    before = _session_count()
    s = open_session(tour_plan_bytes)
    handle = s.handle
    del s
    gc.collect()
    # The finalizer is best-effort; on CPython it should run synchronously here.
    after = _session_count()
    assert after == before, (
        f"finalizer should have closed forgotten session {handle}; "
        f"live count went {before} → {after}"
    )


def test_shutdown_host_releases_everything(tour_plan_bytes: bytes) -> None:
    """``shutdown_host()`` tears the subprocess down cleanly."""
    s = open_session(tour_plan_bytes)
    try:
        assert _session_count() >= 1
    finally:
        s.close()
    shutdown_host()
    # Next call must lazily respawn a fresh host with zero sessions.
    assert _session_count() == 0
