"""Anchor-id persistence across save→reopen (issue #303 ripple).

``DocxSessionSettings.persist_anchor_ids`` (open-time) already flows through the
settings wire; these tests pin it and cover the per-call ``save`` override that
mirrors the MCP server's ``docxodus_save`` ``persistAnchorIds`` argument.
"""

from __future__ import annotations

import io
import zipfile

from docx_scalpel import DocxSession, DocxSessionSettings, Position, open_session


def _first_body_paragraph(session: DocxSession) -> str:
    for anchor in session.project().anchor_index.values():
        if anchor.scope == "body" and anchor.kind in ("p", "h", "li"):
            return anchor.id
    raise AssertionError("fixture has no body paragraph anchors")


def _insert_paragraph(session: DocxSession, markdown: str) -> str:
    """Insert a paragraph and return its created anchor id — a fresh (random) Unid,
    exactly the kind of anchor that cannot survive a save→reopen unless the save
    persists the anchor bookkeeping."""
    result = session.insert_paragraph(_first_body_paragraph(session), Position.AFTER, markdown)
    assert result.success, result.error
    return result.created[0].id


def _document_xml(docx: bytes) -> bytes:
    with zipfile.ZipFile(io.BytesIO(docx)) as z:
        return z.read("word/document.xml")


def test_save_persist_anchor_ids_true_keeps_created_anchor(tour_plan_bytes: bytes) -> None:
    with open_session(tour_plan_bytes) as session:  # default: anchor ids NOT persisted
        created = _insert_paragraph(session, "checkpoint me")
        data = session.save(persist_anchor_ids=True)

    with open_session(data) as reopened:
        result = reopened.replace_text(created, "still addressable")
        assert result.success, result.error


def test_save_persist_anchor_ids_false_strips_on_persist_true_session(
    tour_plan_bytes: bytes,
) -> None:
    settings = DocxSessionSettings(persist_anchor_ids=True)
    with open_session(tour_plan_bytes, settings) as session:
        _insert_paragraph(session, "clean deliverable")
        data = session.save(persist_anchor_ids=False)

    assert b"Unid=" not in _document_xml(data)


def test_open_time_persist_anchor_ids_governs_plain_save(tour_plan_bytes: bytes) -> None:
    settings = DocxSessionSettings(persist_anchor_ids=True)
    with open_session(tour_plan_bytes, settings) as session:
        _insert_paragraph(session, "bookkeeping should survive")
        data = session.save()

    assert b"Unid=" in _document_xml(data)
