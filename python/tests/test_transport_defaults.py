"""Issue #960: per-op defaults, revision filters and result caps are the engine facade's, so the
Python client gets the same behaviour MCP does without restating any default itself."""

from __future__ import annotations

from pathlib import Path

from docx_scalpel import open_session
from docx_scalpel.enums import Position, TrackedChangeMode

FIXTURE = Path(__file__).parents[2] / "TestFiles" / "DA001-TemplateDocument.docx"


def _first_body_anchor(session) -> str:
    return next(a for a in session.project().anchor_index if a.startswith("p:body:"))


def test_table_of_contents_without_position_goes_before_the_anchor():
    with open_session(FIXTURE.read_bytes()) as session:
        anchor = _first_body_anchor(session)
        result = session.insert_table_of_contents(anchor)
        assert result.success, result.error
        ids = list(session.project().anchor_index)
        toc = next(a for a in ids if a.startswith("sdt:body:"))
        assert ids.index(toc) < ids.index(anchor)


def test_list_revisions_filters_by_author_in_the_engine():
    with open_session(FIXTURE.read_bytes()) as session:
        session.set_tracked_changes(TrackedChangeMode.RENDER_INLINE)
        anchor = _first_body_anchor(session)
        session.set_revision_author("Ann")
        assert session.insert_paragraph(anchor, Position.AFTER, "by Ann").success
        session.set_revision_author("Bob")
        assert session.insert_paragraph(anchor, Position.AFTER, "by Bob").success

        everyone = session.list_revisions()
        assert {r.author for r in everyone} >= {"Ann", "Bob"}
        bob = session.list_revisions(author="bob")
        assert bob and all(r.author == "Bob" for r in bob)


def test_grep_max_results_caps_the_matches():
    with open_session(FIXTURE.read_bytes()) as session:
        anchor = _first_body_anchor(session)
        for i in range(3):
            assert session.insert_paragraph(anchor, Position.AFTER, f"needle {i}").success
        assert len(session.grep("needle")) == 3
        assert len(session.grep("needle", max_results=2)) == 2
