"""Which side of tracked deleted text a zero-length span insertion goes on (issue #1290).

Deleted text takes no room in a paragraph's offsets, so the offset just after
``The tenant shall `` names both sides of a deleted ``promptly ``.
``deleted_text_side`` picks one; omitted, the engine keeps its default placement.
"""

from __future__ import annotations

from docx_scalpel import DeletedTextSide, DocxSession, FormatOp, TrackedChangeMode, open_session


def _paragraph(session: DocxSession) -> str:
    for anchor in session.project().anchor_index.values():
        if anchor.scope == "body" and anchor.kind in ("p", "h", "li"):
            return anchor.id
    raise AssertionError("fixture has no body paragraph anchors")


def _tenant_with_deletion(session: DocxSession) -> str:
    anchor = _paragraph(session)
    assert session.replace_text(anchor, "The tenant shall promptly repaint").success
    anchor = _paragraph(session)
    session.set_tracked_changes(TrackedChangeMode.RENDER_INLINE)
    assert session.replace_text_at_span(anchor, 17, 9, "").success
    return _paragraph(session)


def _insertion_follows_deletion(session: DocxSession, anchor: str) -> bool:
    xml = session.raw.get_xml(anchor)
    return xml.index("<w:ins") > xml.index("<w:del ")


def test_default_placement_is_before_the_deletion(tour_plan_bytes: bytes) -> None:
    with open_session(tour_plan_bytes) as session:
        anchor = _tenant_with_deletion(session)
        assert session.replace_text_at_span(anchor, 17, 0, "Y").success
        assert not _insertion_follows_deletion(session, anchor)


def test_after_puts_the_insertion_after_the_deletion(tour_plan_bytes: bytes) -> None:
    with open_session(tour_plan_bytes) as session:
        anchor = _tenant_with_deletion(session)
        assert session.replace_text_at_span(
            anchor, 17, 0, "Y", deleted_text_side=DeletedTextSide.AFTER,
        ).success
        assert _insertion_follows_deletion(session, anchor)


def test_the_side_reaches_the_formatting_variant(tour_plan_bytes: bytes) -> None:
    with open_session(tour_plan_bytes) as session:
        anchor = _tenant_with_deletion(session)
        assert session.replace_text_at_span(
            anchor, 17, 0, "Y", format=FormatOp(bold=True), deleted_text_side=DeletedTextSide.AFTER,
        ).success
        assert _insertion_follows_deletion(session, anchor)
