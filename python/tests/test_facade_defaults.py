"""Issue #1024: an argument the caller leaves out takes the engine facade's default. The Python
client sends nothing for it, rather than restating the default itself, so it cannot drift from the
other transports."""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

from docx_scalpel import TableAnchorEntityKind, open_session
from docx_scalpel import session as session_module
from docx_scalpel.enums import (
    ContextBoundary,
    DiffFormat,
    PlaceholderKinds,
    Position,
    ProjectionDepth,
    ProjectionScopes,
    RegexOptions,
)

FIXTURE = Path(__file__).parents[2] / "TestFiles" / "DA001-TemplateDocument.docx"


def _first_body_anchor(session) -> str:
    return next(a for a in session.project().anchor_index if a.startswith("p:body:"))


def _record(monkeypatch: pytest.MonkeyPatch) -> list[tuple[str, dict[str, Any]]]:
    """Record each request a session sends, then send it on to the host."""
    sent: list[tuple[str, dict[str, Any]]] = []
    real = session_module._call

    def call(op: str, payload: dict[str, Any]) -> Any:
        sent.append((op, dict(payload)))
        return real(op, payload)

    monkeypatch.setattr(session_module, "_call", call)
    return sent


def test_omitted_arguments_are_left_to_the_engine(monkeypatch: pytest.MonkeyPatch):
    with open_session(FIXTURE.read_bytes()) as session:
        anchor = _first_body_anchor(session)
        sent = _record(monkeypatch)

        assert session.get_diff() == session.get_diff(DiffFormat.JSON)
        session.project_anchor(anchor)
        session.find_by_regex("the")
        session.find_placeholders()
        session.remaining_placeholders()
        assert session.insert_page_number_field(anchor).success
        assert session.set_page_numbering(anchor).success
        assert session.set_page_setup(anchor).success

        defaulted = {
            "get_diff": "format",
            "project_anchor": "depth",
            "find_by_regex": "regexOptions",
            "find_placeholders": "kinds",
            "remaining_placeholders": "kinds",
            "insert_page_number_field": "field",
            "set_page_numbering": "op",
            "set_page_setup": "op",
        }
        # sent[1] is the diff that names its format, to compare with the one that omits it.
        for op, args in sent[:1] + sent[2:]:
            if op in defaulted:
                assert defaulted[op] not in args, f"{op} sent its own {defaulted[op]}"
        assert {op for op, _ in sent} >= set(defaulted)
        placeholders = next(args for op, args in sent if op == "find_placeholders")
        assert not {"kinds", "scope", "boundary"} & placeholders.keys()


def test_omitted_arguments_answer_as_the_engine_default_does():
    with open_session(FIXTURE.read_bytes()) as session:
        anchor = _first_body_anchor(session)

        assert session.project_anchor(anchor) == session.project_anchor(
            anchor, ProjectionDepth.SUBTREE_AND_FOLLOWING_SIBLINGS)
        assert session.find_by_regex("the") == session.find_by_regex("the", RegexOptions.NONE)
        assert session.find_placeholders() == session.find_placeholders(
            PlaceholderKinds.ALL, ProjectionScopes.BODY, None, ContextBoundary.CHAR)
        assert session.remaining_placeholders() == session.remaining_placeholders(PlaceholderKinds.ALL)


def test_set_repeat_header_row_without_repeat_marks_the_row(monkeypatch: pytest.MonkeyPatch):
    with open_session(FIXTURE.read_bytes()) as session:
        anchor = _first_body_anchor(session)
        inserted = session.insert_table(anchor, Position.AFTER, 2, 2)
        assert inserted.success, inserted.error
        cell = inserted.created[0].id
        table = next(
            location.anchor.id for location in inserted.table_anchors.added
            if location.entity_kind is TableAnchorEntityKind.TABLE
        )
        sent = _record(monkeypatch)

        assert session.set_repeat_header_row(cell).success
        assert "repeat" not in sent[-1][1]
        assert "tblHeader" in session.raw.get_xml(table)

        assert session.set_repeat_header_row(cell, False).success
        assert "tblHeader" not in session.raw.get_xml(table)
