"""Issue #963: a session whose failed edit could not be rolled back reports ``session_corrupted``.

``EditErrorCode._missing_`` degrades an unknown wire code to ``INTERNAL_ERROR``, so a client that
never declared the member would silently report the corruption as an ordinary internal error.
Decoding the wire string and demanding the exact member catches that.
"""

from __future__ import annotations

from docx_scalpel.enums import EditErrorCode


def test_session_corrupted_decodes_from_its_wire_string() -> None:
    assert EditErrorCode.SESSION_CORRUPTED.value == "session_corrupted"
    assert EditErrorCode("session_corrupted") is EditErrorCode.SESSION_CORRUPTED
