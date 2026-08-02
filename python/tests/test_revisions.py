"""Selective per-revision accept/reject (issue #318).

``list_revisions`` reads tracked revisions directly off the live markup with
stable ids and the markup's true authors; ``accept_revision``/``reject_revision``
resolve one revision at a time as undoable session mutations.
"""

from __future__ import annotations

import io
import zipfile

from docx_scalpel import DocxSession, TrackedChangeMode, open_session


def _first_body_paragraph(session: DocxSession) -> str:
    for anchor in session.project().anchor_index.values():
        if anchor.scope == "body" and anchor.kind in ("p", "h", "li"):
            return anchor.id
    raise AssertionError("fixture has no body paragraph anchors")


def _document_xml(docx: bytes) -> bytes:
    with zipfile.ZipFile(io.BytesIO(docx)) as z:
        return z.read("word/document.xml")


def test_list_revisions_reads_markup_identity(tour_plan_bytes: bytes) -> None:
    with open_session(tour_plan_bytes) as session:
        session.set_tracked_changes(TrackedChangeMode.RENDER_INLINE)
        session.set_revision_author("py-reviewer")
        anchor = _first_body_paragraph(session)
        assert session.replace_text(anchor, "Tracked rewrite.").success

        revisions = session.list_revisions()
        assert len(revisions) == 2  # one delete (old text) + one insert (new text)
        assert {r.type for r in revisions} == {"delete", "insert"}
        assert all(r.author == "py-reviewer" for r in revisions)
        assert all(r.id.startswith("rev") for r in revisions)
        insert = next(r for r in revisions if r.type == "insert")
        assert insert.text == "Tracked rewrite."
        assert insert.anchor_id is not None


def test_accept_and_reject_one_revision_each(tour_plan_bytes: bytes) -> None:
    with open_session(tour_plan_bytes) as session:
        session.set_tracked_changes(TrackedChangeMode.RENDER_INLINE)
        anchor = _first_body_paragraph(session)
        assert session.replace_text(anchor, "Selective resolution.").success

        revisions = session.list_revisions()
        insert = next(r for r in revisions if r.type == "insert")
        delete = next(r for r in revisions if r.type == "delete")

        accepted = session.accept_revision(insert.id)
        assert accepted.success, accepted.error
        assert accepted.modified

        # The other revision's id is untouched by resolving the first.
        remaining = session.list_revisions()
        assert [r.id for r in remaining] == [delete.id]

        assert session.accept_revision(delete.id).success
        assert session.list_revisions() == ()

        xml = _document_xml(session.save())
        assert b"Selective resolution." in xml
        assert b"w:ins" not in xml
        assert b"w:del" not in xml


def test_unknown_revision_id_fails_with_revision_not_found(
    tour_plan_bytes: bytes,
) -> None:
    with open_session(tour_plan_bytes) as session:
        result = session.accept_revision("rev999999")
        assert not result.success
        assert result.error is not None
        assert result.error.code == "revision_not_found"


def test_reject_revision_is_undoable(tour_plan_bytes: bytes) -> None:
    with open_session(tour_plan_bytes) as session:
        session.set_tracked_changes(TrackedChangeMode.RENDER_INLINE)
        anchor = _first_body_paragraph(session)
        assert session.replace_text(anchor, "Rejected then undone.").success

        insert = next(r for r in session.list_revisions() if r.type == "insert")
        assert session.reject_revision(insert.id).success
        assert insert.id not in [r.id for r in session.list_revisions()]

        assert session.undo()
        assert insert.id in [r.id for r in session.list_revisions()]
