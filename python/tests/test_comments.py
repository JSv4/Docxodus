"""Native Word comment authoring through the Python wrapper (issue #300).

Exercises ``DocxSession.add_comment`` / ``update_comment`` / ``remove_comment`` /
``list_comments`` end-to-end over the stdio host: creation, the created-anchor
contract, listing, body update with attribute preservation, removal, save/reopen
survival, and the typed error envelope.
"""

from __future__ import annotations

from typing import Iterator

import pytest

from docx_scalpel import CharSpan, DocxSession, open_session
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


def test_add_comment_creates_and_projects_it(session: DocxSession) -> None:
    host = _first_body_paragraph(session)

    result = session.add_comment(
        host, None, "Alice", "Needs review.", initials="AL", date="2026-08-01T00:00:00Z"
    )

    assert result.success, result.error
    kinds = {(a.kind, a.scope) for a in result.created}
    assert ("cmt", "cmt") in kinds
    assert ("p", "cmt") in kinds
    assert any(a.id == host for a in result.modified)

    markdown = session.project().markdown
    assert "# Comments" in markdown
    assert "Needs review." in markdown


def test_list_comments_returns_metadata(session: DocxSession) -> None:
    host = _first_body_paragraph(session)
    made = session.add_comment(
        host, None, "Alice", "First body.", initials="AL", date="2026-08-01T00:00:00Z"
    )
    assert made.success, made.error
    assert session.add_comment(host, None, "Bob", "Second body.").success

    entries = session.list_comments()
    assert len(entries) == 2
    assert entries[0].author == "Alice"
    assert entries[0].initials == "AL"
    assert entries[0].date == "2026-08-01T00:00:00Z"
    assert entries[0].text == "First body."
    assert entries[0].anchor_id == next(a.id for a in made.created if a.kind == "cmt")
    assert entries[1].author == "Bob"
    assert entries[1].initials is None
    assert entries[1].date is None


def test_update_comment_replaces_body_and_preserves_author(session: DocxSession) -> None:
    host = _first_body_paragraph(session)
    made = session.add_comment(host, CharSpan(0, 4), "Alice", "Original.", initials="AL")
    assert made.success, made.error
    cmt_anchor = next(a.id for a in made.created if a.kind == "cmt")

    updated = session.update_comment(cmt_anchor, "Revised body.")
    assert updated.success, updated.error

    entry = session.list_comments()[0]
    assert entry.text == "Revised body."
    assert entry.author == "Alice"
    assert entry.initials == "AL"


def test_remove_comment_and_save_reopen_round_trip(session: DocxSession) -> None:
    host = _first_body_paragraph(session)
    keep = session.add_comment(host, None, "Alice", "Keep me.")
    drop = session.add_comment(host, None, "Bob", "Drop me.")
    assert keep.success and drop.success

    removed = session.remove_comment(next(a.id for a in drop.created if a.kind == "cmt"))
    assert removed.success, removed.error
    assert [e.text for e in session.list_comments()] == ["Keep me."]

    reopened = open_session(session.save())
    try:
        assert [e.text for e in reopened.list_comments()] == ["Keep me."]
        assert "Keep me." in reopened.project().markdown
    finally:
        reopened.close()


def test_typed_errors(session: DocxSession) -> None:
    host = _first_body_paragraph(session)

    empty_span = session.add_comment(host, CharSpan(0, 0), "A", "x")
    assert not empty_span.success
    assert empty_span.error is not None
    assert empty_span.error.code is EditErrorCode.EMPTY_COMMENT_SPAN

    made = session.add_comment(host, None, "A", "Host.")
    assert made.success, made.error
    cmt_para = next(a.id for a in made.created if a.kind == "p" and a.scope == "cmt")

    # Word has no comments-on-comments: a cmt-scope paragraph is not a legal host.
    nested = session.add_comment(cmt_para, None, "A", "Nested.")
    assert not nested.success
    assert nested.error is not None
    assert nested.error.code is EditErrorCode.ANCHOR_WRONG_KIND

    wrong_kind = session.update_comment(host, "Nope.")
    assert not wrong_kind.success
    assert wrong_kind.error is not None
    assert wrong_kind.error.code is EditErrorCode.ANCHOR_WRONG_KIND
