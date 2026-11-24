"""The document outline (Word's Navigation Pane): ``list_outline`` and ``move_block_range``."""

from __future__ import annotations

from pathlib import Path

import pytest

from docx_scalpel import Position, open_session

FIXTURE = Path(__file__).parents[2] / "TestFiles" / "DA001-TemplateDocument.docx"


def _session():
    if not FIXTURE.exists():
        pytest.skip(f"fixture missing: {FIXTURE}")
    session = open_session(FIXTURE.read_bytes())
    first = next(a for a in session.project().anchor_index.values() if a.scope == "body")
    result = session.insert_paragraph(
        first.id, Position.BEFORE, "Alpha\n\nalpha text\n\nAlpha sub\n\nsub text\n\nBeta\n\nbeta text",
    )
    assert result.success, result.error
    # Set the heading styles by id (the session defines Heading1/Heading2 when the document lacks them).
    for index, style in ((0, "Heading1"), (2, "Heading2"), (4, "Heading1")):
        assert session.set_paragraph_style(result.created[index].id, style).success
    return session


def _entries(session):
    return {entry.text: entry for entry in session.list_outline()}


def test_list_outline_reports_levels_and_each_heading_part() -> None:
    with _session() as session:
        entries = _entries(session)
        alpha, sub, beta = entries["Alpha"], entries["Alpha sub"], entries["Beta"]
        assert (alpha.level, sub.level, beta.level) == (1, 2, 1)
        assert alpha.style_id is not None
        order = [entry.text for entry in session.list_outline()]
        assert order.index("Alpha") < order.index("Alpha sub") < order.index("Beta")
        # Alpha's part runs up to Beta: its last block is the subheading's text.
        assert alpha.last_anchor_id == sub.last_anchor_id
        assert alpha.last_anchor_id != beta.anchor_id


def test_move_block_range_moves_a_heading_with_its_text_in_one_undo() -> None:
    with _session() as session:
        entries = _entries(session)
        alpha, beta = entries["Alpha"], entries["Beta"]

        result = session.move_block_range(alpha.anchor_id, alpha.last_anchor_id, beta.last_anchor_id, Position.AFTER)

        assert result.success, result.error
        order = [entry.text for entry in session.list_outline()]
        assert order.index("Beta") < order.index("Alpha") < order.index("Alpha sub")
        assert session.undo()
        order = [entry.text for entry in session.list_outline()]
        assert order.index("Alpha") < order.index("Beta")


def test_move_block_range_onto_its_own_run_is_refused() -> None:
    with _session() as session:
        alpha = _entries(session)["Alpha"]
        sub = _entries(session)["Alpha sub"]
        refused = session.move_block_range(alpha.anchor_id, alpha.last_anchor_id, sub.anchor_id, Position.AFTER)
        assert not refused.success
