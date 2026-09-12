"""Nested-fill policies, child fills and the per-operation matrix through the stdio host (issue #763).

The fixture is ``TestFiles/CC763-NestedContentControls.docx``: an outer block rich-text
control (native id 100, tag ``outer-tag``) whose paragraph holds an inline plain-text
child (native id 101), plus a checkbox (102).
"""

from __future__ import annotations

from pathlib import Path

import pytest

from docx_scalpel import (
    ContentControlFillOptions,
    ContentControlNestedPolicy,
    DocxSession,
    DocxSessionSettings,
    EditErrorCode,
    TrackedChangeMode,
    open_session,
)


@pytest.fixture(scope="session")
def nested_bytes(test_files_dir: Path) -> bytes:
    return (test_files_dir / "CC763-NestedContentControls.docx").read_bytes()


def _control(session: DocxSession, native_id: str):
    return next(c for c in session.list_content_controls() if c.native_id == native_id)


def test_preserve_keeps_the_nested_child_and_fills_it_by_anchor(nested_bytes: bytes) -> None:
    with open_session(nested_bytes) as session:
        outer = _control(session, "100")
        inner = _control(session, "101")
        assert outer.nested_control_anchor_ids == (inner.anchor_id,)
        preserve = next(op for op in outer.operations
                        if op.operation == "fill_text" and op.nested_controls is ContentControlNestedPolicy.PRESERVE)
        assert preserve.can_mutate
        refuse = next(op for op in outer.operations
                      if op.operation == "fill_text" and op.nested_controls is ContentControlNestedPolicy.REFUSE)
        assert not refuse.can_mutate and refuse.reason

        refused = session.fill_content_control_text(outer.anchor_id, "x")
        assert refused.error is not None
        assert refused.error.code is EditErrorCode.CONTENT_CONTROL_NESTED_FILL_UNSUPPORTED

        result = session.fill_content_control_text(
            outer.anchor_id,
            "Outer via python",
            ContentControlFillOptions(
                nested_controls=ContentControlNestedPolicy.PRESERVE,
                child_fills={inner.anchor_id: "Inner via python"},
            ),
        )
        assert result.success, result.error
        assert [a.id for a in result.modified] == [outer.anchor_id, inner.anchor_id]
        assert _control(session, "101").text == "Inner via python"
        assert "Outer via python" in _control(session, "100").text
        assert _control(session, "100").tag == "outer-tag"


def test_replace_reports_the_dropped_child(nested_bytes: bytes) -> None:
    with open_session(nested_bytes) as session:
        outer = _control(session, "100")
        inner = _control(session, "101")
        result = session.fill_content_control_text(
            outer.anchor_id, "flat",
            ContentControlFillOptions(nested_controls=ContentControlNestedPolicy.REPLACE),
        )
        assert result.success, result.error
        assert [a.id for a in result.removed] == [inner.anchor_id]
        assert _control(session, "100").text == "flat"
        assert all(c.native_id != "101" for c in session.list_content_controls())


def test_tracked_text_fill_records_revisions_and_state_changes_explain_themselves(nested_bytes: bytes) -> None:
    settings = DocxSessionSettings(tracked_changes=TrackedChangeMode.RENDER_INLINE)
    with open_session(nested_bytes, settings) as session:
        inner = _control(session, "101")
        assert inner.can_mutate
        checkbox = _control(session, "102")
        assert not checkbox.can_mutate
        assert checkbox.unsupported_reason and "w14:checked" in checkbox.unsupported_reason

        result = session.fill_content_control_text(inner.anchor_id, "tracked via python")
        assert result.success, result.error
        types = {r.type for r in session.list_revisions()}
        assert {"insert", "delete"} <= types
        refused = session.set_content_control_checked(checkbox.anchor_id, True)
        assert refused.error is not None
        assert refused.error.code is EditErrorCode.TRACKED_OPERATION_UNSUPPORTED
