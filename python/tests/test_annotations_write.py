"""End-to-end tests for the annotation write surface through docx-scalpel."""

from __future__ import annotations

from pathlib import Path

import pytest

from docx_scalpel import (
    AnnotationUpdate,
    CharSpan,
    DocumentAnnotation,
    open_session,
)
from docx_scalpel.enums import EditErrorCode

FIXTURE = Path(__file__).parents[2] / "TestFiles" / "DA001-TemplateDocument.docx"


def _first_paragraph_anchor(session) -> str:
    projection = session.project()
    for anchor_id, info in projection.anchor_index.items():
        if anchor_id.startswith("p:body:") and info.text_preview:
            return anchor_id
    pytest.fail("no body paragraph with text in fixture")


def test_add_and_remove_round_trip():
    with open_session(FIXTURE.read_bytes()) as session:
        anchor = _first_paragraph_anchor(session)
        r = session.add_annotation(
            anchor,
            CharSpan(start=0, length=1),
            DocumentAnnotation(
                id="py-1", label_id="LBL", label="Lbl",
                color="#FF0", bookmark_name="",
            ),
        )
        assert r.success
        assert r.annotation_id == "py-1"
        assert any(a.id == "py-1" for a in session.list_annotations())

        r = session.remove_annotation("py-1")
        assert r.success
        assert not any(a.id == "py-1" for a in session.list_annotations())


def test_add_with_auto_id():
    with open_session(FIXTURE.read_bytes()) as session:
        anchor = _first_paragraph_anchor(session)
        r = session.add_annotation(
            anchor, CharSpan(0, 1),
            DocumentAnnotation(id="", label_id="L", label="L", color="#000", bookmark_name=""),
        )
        assert r.success
        assert r.annotation_id is not None
        assert len(r.annotation_id) == 16


def test_add_anchor_not_found_returns_error():
    with open_session(FIXTURE.read_bytes()) as session:
        r = session.add_annotation(
            "p:body:DEADBEEFDEADBEEF", CharSpan(0, 1),
            DocumentAnnotation(id="x", label_id="L", label="L", color="#000", bookmark_name=""),
        )
        assert not r.success
        assert r.error is not None
        assert r.error.code == EditErrorCode.ANCHOR_NOT_FOUND


def test_update_mutates_metadata():
    with open_session(FIXTURE.read_bytes()) as session:
        anchor = _first_paragraph_anchor(session)
        session.add_annotation(
            anchor, CharSpan(0, 1),
            DocumentAnnotation(
                id="py-u", label_id="L", label="L", color="#000", bookmark_name="",
                metadata={"keep": "yes", "drop": "old"},
            ),
        )
        r = session.update_annotation(
            "py-u",
            AnnotationUpdate(label="New", metadata_patch={"drop": None, "new": "fresh"}),
        )
        assert r.success
        listed = next(a for a in session.list_annotations() if a.id == "py-u")
        assert listed.label == "New"
        assert listed.metadata == {"keep": "yes", "new": "fresh"}


def test_move_retargets_modified_includes_both_blocks():
    with open_session(FIXTURE.read_bytes()) as session:
        projection = session.project()
        anchors = [
            a for a, info in projection.anchor_index.items()
            if a.startswith("p:body:") and info.text_preview
        ][:2]
        assert len(anchors) == 2

        session.add_annotation(
            anchors[0], CharSpan(0, 1),
            DocumentAnnotation(id="py-m", label_id="L", label="L", color="#000", bookmark_name=""),
        )
        r = session.move_annotation("py-m", anchors[1], CharSpan(0, 2))
        assert r.success
        modified_ids = {m.id for m in r.modified}
        assert anchors[0] in modified_ids
        assert anchors[1] in modified_ids


def test_save_reopen_persists_annotation():
    with open_session(FIXTURE.read_bytes()) as session:
        anchor = _first_paragraph_anchor(session)
        session.add_annotation(
            anchor, CharSpan(0, 4),
            DocumentAnnotation(
                id="persist", label_id="P", label="P", color="#0F0", bookmark_name="",
                metadata={"k": "v"},
            ),
        )
        saved = session.save()

    with open_session(saved) as reopened:
        listed = next(a for a in reopened.list_annotations() if a.id == "persist")
        assert listed.label_id == "P"
        assert listed.metadata["k"] == "v"
