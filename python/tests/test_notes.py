"""Footnote / endnote authoring through the Python wrapper (issue #276).

Exercises ``DocxSession.insert_footnote`` / ``insert_endnote`` end-to-end over the
stdio host: creation, the created-anchor contract, editing and deleting an authored
note through the pre-existing text ops, and the typed error envelope.
"""

from __future__ import annotations

from typing import Iterator

import pytest

from docx_scalpel import DocxSession, open_session
from docx_scalpel.enums import EditErrorCode


@pytest.fixture
def session(tour_plan_bytes: bytes) -> Iterator[DocxSession]:
    s = open_session(tour_plan_bytes)
    try:
        yield s
    finally:
        s.close()


def _first_body_paragraph(session: DocxSession) -> str:
    for anchor in session.project().anchor_index.values():
        if anchor.scope == "body" and anchor.kind in ("p", "h", "li"):
            return anchor.id
    pytest.skip("fixture has no body paragraph anchors")


def test_insert_footnote_creates_note_and_projects_it(session: DocxSession) -> None:
    host = _first_body_paragraph(session)

    result = session.insert_footnote(host, 0, "Source: 2025 annual report.")

    assert result.success, result.error
    kinds = {(a.kind, a.scope) for a in result.created}
    assert ("fn", "fn") in kinds
    assert ("p", "fn") in kinds
    assert any(a.id == host for a in result.modified)

    markdown = session.project().markdown
    assert "# Footnotes" in markdown
    assert "Source: 2025 annual report." in markdown


def test_insert_endnote_uses_the_endnote_scope(session: DocxSession) -> None:
    host = _first_body_paragraph(session)

    result = session.insert_endnote(host, 0, "See appendix B.")

    assert result.success, result.error
    assert any(a.kind == "en" for a in result.created)
    assert "See appendix B." in session.project().markdown


def test_authored_footnote_is_editable_and_deletable(session: DocxSession) -> None:
    host = _first_body_paragraph(session)
    created = session.insert_footnote(host, 0, "Original.")
    assert created.success, created.error

    note_para = next(a.id for a in created.created if a.kind == "p" and a.scope == "fn")
    note_def = next(a.id for a in created.created if a.kind == "fn")

    edited = session.replace_text(note_para, "Rewritten.")
    assert edited.success, edited.error
    assert "Rewritten." in session.project().markdown

    deleted = session.delete_block(note_def)
    assert deleted.success, deleted.error
    assert "Rewritten." not in session.project().markdown


def test_note_survives_a_save_and_reopen(session: DocxSession) -> None:
    host = _first_body_paragraph(session)
    assert session.insert_footnote(host, 0, "Persisted note.").success

    reopened = open_session(session.save())
    try:
        assert "Persisted note." in reopened.project().markdown
    finally:
        reopened.close()


def test_typed_errors_for_bad_host_and_offset(session: DocxSession) -> None:
    host = _first_body_paragraph(session)
    created = session.insert_footnote(host, 0, "A note.")
    assert created.success, created.error
    note_para = next(a.id for a in created.created if a.kind == "p" and a.scope == "fn")

    # Word does not allow a note reference inside another note's story.
    nested = session.insert_footnote(note_para, 0, "Nested.")
    assert not nested.success
    assert nested.error is not None
    assert nested.error.code is EditErrorCode.ANCHOR_WRONG_KIND

    out_of_range = session.insert_endnote(host, 100_000, "Nope.")
    assert not out_of_range.success
    assert out_of_range.error is not None
    assert out_of_range.error.code is EditErrorCode.OFFSET_OUT_OF_RANGE
