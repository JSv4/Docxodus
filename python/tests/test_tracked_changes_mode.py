"""Mid-session tracked-changes mode switching (issue #304).

``DocxSessionSettings.tracked_changes`` used to be fixed for a session's lifetime;
these tests pin the ``set_tracked_changes`` / ``set_revision_author`` mutators that
switch recording mid-workflow without a close+reopen.
"""

from __future__ import annotations

import io
import zipfile

from docx_scalpel import DocxSession, TrackedChangeMode, open_session


def _body_paragraphs(session: DocxSession, count: int) -> list[str]:
    ids = [
        anchor.id
        for anchor in session.project().anchor_index.values()
        if anchor.scope == "body" and anchor.kind in ("p", "h", "li")
    ]
    assert len(ids) >= count, "fixture has too few body paragraph anchors"
    return ids[:count]


def _first_body_paragraph(session: DocxSession) -> str:
    return _body_paragraphs(session, 1)[0]


def _document_xml(docx: bytes) -> bytes:
    with zipfile.ZipFile(io.BytesIO(docx)) as z:
        return z.read("word/document.xml")


def test_set_tracked_changes_mid_session(tour_plan_bytes: bytes) -> None:
    with open_session(tour_plan_bytes) as session:  # default: accept (direct edits)
        anchor = _first_body_paragraph(session)

        direct = session.replace_text(anchor, "Direct edit.")
        assert direct.success, direct.error
        assert b"w:ins" not in _document_xml(session.save())

        session.set_tracked_changes(TrackedChangeMode.RENDER_INLINE)
        session.set_revision_author("py-reviewer")

        tracked = session.replace_text(anchor, "Tracked edit.")
        assert tracked.success, tracked.error

        xml = _document_xml(session.save())
        assert b"w:ins" in xml
        assert b"py-reviewer" in xml


def test_switch_back_to_accept_leaves_history(tour_plan_bytes: bytes) -> None:
    with open_session(tour_plan_bytes) as session:
        # Distinct anchors: a direct edit on a paragraph REWRITES that paragraph's
        # content, revision markup included, so the untouched-history check must
        # target a paragraph the second edit never touches (mirrors DS331).
        tracked_anchor, direct_anchor = _body_paragraphs(session, 2)

        session.set_tracked_changes(TrackedChangeMode.RENDER_INLINE)
        tracked = session.replace_text(tracked_anchor, "Tracked edit.")
        assert tracked.success, tracked.error
        ins_count = _document_xml(session.save()).count(b"<w:ins ")
        assert ins_count > 0

        session.set_tracked_changes(TrackedChangeMode.ACCEPT)
        direct = session.replace_text(direct_anchor, "Direct edit.")
        assert direct.success, direct.error

        xml = _document_xml(session.save())
        assert xml.count(b"<w:ins ") == ins_count  # history untouched
        assert b"Direct edit." in xml
