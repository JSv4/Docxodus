"""Session ops the stdio host and the Python client used to leave out (issue #1026).

``insert_horizontal_rule`` and ``list_notes`` reached every other transport but not this
one. Each test calls the method over the real stdio host and checks the effect on the
document, not just a success flag.
"""

from __future__ import annotations

from typing import Iterator

import pytest

from docx_scalpel import DocxSession, NoteListEntry, ParagraphBorderEdge, open_session
from docx_scalpel.enums import Position


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


def test_insert_horizontal_rule_inserts_a_bordered_paragraph(session: DocxSession) -> None:
    host = _first_body_paragraph(session)

    result = session.insert_horizontal_rule(
        host, Position.AFTER, ParagraphBorderEdge(style="double", size=4)
    )

    assert result.success, result.error
    rule = result.created[0].id
    xml = session.raw.get_xml(rule)
    assert "pBdr" in xml
    assert 'w:val="double"' in xml
    assert 'w:sz="4"' in xml


def test_insert_horizontal_rule_without_a_rule_takes_the_default_edge(session: DocxSession) -> None:
    host = _first_body_paragraph(session)

    result = session.insert_horizontal_rule(host, Position.BEFORE)

    assert result.success, result.error
    xml = session.raw.get_xml(result.created[0].id)
    assert "pBdr" in xml
    assert 'w:val="single"' in xml


def test_list_notes_lists_footnotes_and_endnotes_in_citation_order(session: DocxSession) -> None:
    host = _first_body_paragraph(session)
    first = session.insert_footnote(host, 0, "First note.")
    second = session.insert_footnote(host, 0, "Earlier note.")
    endnote = session.insert_endnote(host, 0, "An endnote.")
    assert first.success and second.success and endnote.success

    footnotes = session.list_notes(endnotes=False)
    endnotes = session.list_notes(endnotes=True)

    assert all(isinstance(n, NoteListEntry) for n in footnotes)
    assert [n.ordinal for n in footnotes] == [1, 2]
    second_def = next(a.id for a in second.created if a.kind == "fn")
    first_def = next(a.id for a in first.created if a.kind == "fn")
    # Both cite offset 0, so the later insertion sits before the first and is cited first.
    assert [n.def_anchor_id for n in footnotes] == [second_def, first_def]
    assert [n.def_anchor_id for n in endnotes] == [next(a.id for a in endnote.created if a.kind == "en")]
