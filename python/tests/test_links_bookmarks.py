"""Native hyperlinks and bookmarks through the Python wrapper (issue #451).

Peer suite to ``Docxodus.Tests/DocxSessionLinkBookmarkTests.cs``: exercises the
hyperlink and bookmark surface end-to-end over the stdio host, including inbound
reference retargeting, the reserved-name policy, save/reopen survival, and the
typed errors an agent is expected to branch on.
"""

from __future__ import annotations

from typing import Iterator

import pytest

from docx_scalpel import (
    CharSpan,
    DocumentRange,
    DocxSession,
    HeaderFooterKind,
    HyperlinkKind,
    ProjectionScopes,
    TrackedChangeMode,
    open_session,
)
from docx_scalpel.enums import EditErrorCode

# Two paragraphs are normalized to known text so every span in this module is a
# literal offset rather than a guess about the fixture's wording.
FIRST_TEXT = "Alpha beta gamma delta"
SECOND_TEXT = "Second paragraph text"


@pytest.fixture
def session(tour_plan_bytes: bytes) -> Iterator[DocxSession]:
    s = open_session(tour_plan_bytes)
    try:
        yield s
    finally:
        s.close()


def _body_paragraphs(session: DocxSession) -> list[str]:
    seen: list[str] = []
    for anchor in session.project().anchor_index.values():
        if anchor.scope == "body" and anchor.kind in ("p", "h", "li") and anchor.id not in seen:
            seen.append(anchor.id)
    if len(seen) < 2:
        pytest.skip("fixture has fewer than two body paragraph anchors")
    return seen


@pytest.fixture
def paragraphs(session: DocxSession) -> list[str]:
    anchors = _body_paragraphs(session)
    assert session.replace_text(anchors[0], FIRST_TEXT).success
    assert session.replace_text(anchors[1], SECOND_TEXT).success
    return anchors


def test_external_hyperlink_crud_and_relationship_reuse(
    session: DocxSession, paragraphs: list[str]
) -> None:
    first = session.add_hyperlink(
        paragraphs[0], CharSpan(0, 5), HyperlinkKind.EXTERNAL, "https://example.test/shared"
    )
    second = session.add_hyperlink(
        paragraphs[1], CharSpan(0, 6), HyperlinkKind.EXTERNAL, "https://example.test/shared"
    )
    assert first.success, first.error
    assert second.success, second.error

    links = session.list_hyperlinks()
    assert len(links) == 2
    assert {link.kind for link in links} == {HyperlinkKind.EXTERNAL}
    assert links[0].text == "Alpha"
    # One URI, one owning-part relationship, reused by both links.
    assert len({link.relationship_id for link in links}) == 1
    assert not any(link.is_broken for link in links)

    updated = session.update_hyperlink(
        first.hyperlink_id, HyperlinkKind.EXTERNAL, "https://example.test/moved"
    )
    assert updated.success, updated.error
    targets = {link.id: link.target for link in session.list_hyperlinks()}
    assert targets[first.hyperlink_id] == "https://example.test/moved"
    assert targets[second.hyperlink_id] == "https://example.test/shared"

    assert session.remove_hyperlink(first.hyperlink_id).success
    surviving = session.list_hyperlinks()
    assert [link.id for link in surviving] == [second.hyperlink_id]

    # Identity survives a save/reopen round trip.
    with open_session(session.save(persist_anchor_ids=True)) as reopened:
        assert [link.id for link in reopened.list_hyperlinks()] == [second.hyperlink_id]


def test_bookmark_range_reports_segments_and_moves(
    session: DocxSession, paragraphs: list[str]
) -> None:
    added = session.add_bookmark(
        "AcrossParas", DocumentRange(paragraphs[0], 6, paragraphs[1], 6)
    )
    assert added.success, added.error
    assert added.bookmark_name == "AcrossParas"

    bookmark = next(b for b in session.list_bookmarks() if b.name == "AcrossParas")
    assert bookmark.is_valid, bookmark.validation_error
    assert bookmark.is_paired
    assert not bookmark.is_managed
    assert len(bookmark.segments) == 2
    assert bookmark.text == "beta gamma delta\nSecond"
    assert bookmark.range == DocumentRange(paragraphs[0], 6, paragraphs[1], 6)

    moved = session.move_bookmark(
        "AcrossParas", DocumentRange(paragraphs[1], 0, paragraphs[1], 6)
    )
    assert moved.success, moved.error
    after = next(b for b in session.list_bookmarks() if b.name == "AcrossParas")
    # A same-part move keeps the pair's numeric id; only the coordinates change.
    assert after.bookmark_id == bookmark.bookmark_id
    assert after.text == "Second"


def test_internal_link_rename_retargets_and_removal_is_refused(
    session: DocxSession, paragraphs: list[str]
) -> None:
    assert session.add_bookmark(
        "TargetOne", DocumentRange(paragraphs[0], 0, paragraphs[0], 5)
    ).success
    link = session.add_hyperlink(
        paragraphs[1], CharSpan(0, 6), HyperlinkKind.INTERNAL, "TargetOne"
    )
    assert link.success, link.error
    internal = next(x for x in session.list_hyperlinks() if x.id == link.hyperlink_id)
    assert internal.kind is HyperlinkKind.INTERNAL
    assert internal.target == "TargetOne"
    # Internal links are relationship-free w:anchor markup.
    assert internal.relationship_id is None
    assert not internal.is_broken

    blocked = session.remove_bookmark("TargetOne")
    assert not blocked.success
    assert blocked.error is not None
    assert blocked.error.code is EditErrorCode.BOOKMARK_IN_USE

    renamed = session.rename_bookmark("TargetOne", "TargetTwo")
    assert renamed.success, renamed.error
    retargeted = next(x for x in session.list_hyperlinks() if x.id == link.hyperlink_id)
    assert retargeted.target == "TargetTwo"
    assert not retargeted.is_broken
    assert {b.name for b in session.list_bookmarks()} >= {"TargetTwo"}

    # Releasing the last inbound reference releases the bookmark.
    assert session.update_hyperlink(
        link.hyperlink_id, HyperlinkKind.EXTERNAL, "https://example.test/out"
    ).success
    assert session.remove_bookmark("TargetTwo").success
    assert all(b.name != "TargetTwo" for b in session.list_bookmarks())


def test_reserved_word_bookmark_namespace_is_closed_to_creation(
    session: DocxSession, paragraphs: list[str]
) -> None:
    span = DocumentRange(paragraphs[0], 0, paragraphs[0], 5)
    for reserved in ("_GoBack", "_Toc12345", "_Ref99", "_Hlt7", "_Hlk7"):
        created = session.add_bookmark(reserved, span)
        assert not created.success, reserved
        assert created.error is not None
        assert created.error.code is EditErrorCode.INVALID_BOOKMARK_NAME

    assert session.add_bookmark("Ordinary", span).success
    renamed = session.rename_bookmark("Ordinary", "_Toc1")
    assert not renamed.success
    assert renamed.error is not None
    assert renamed.error.code is EditErrorCode.INVALID_BOOKMARK_NAME


def test_structured_errors_for_missing_targets_and_bad_spans(
    session: DocxSession, paragraphs: list[str]
) -> None:
    missing = session.add_hyperlink(
        paragraphs[0], CharSpan(0, 5), HyperlinkKind.INTERNAL, "NoSuchBookmark"
    )
    assert not missing.success
    assert missing.error is not None
    assert missing.error.code is EditErrorCode.MISSING_BOOKMARK_TARGET

    empty = session.add_hyperlink(
        paragraphs[0], CharSpan(0, 0), HyperlinkKind.EXTERNAL, "https://example.test/x"
    )
    assert not empty.success
    assert empty.error is not None
    assert empty.error.code is EditErrorCode.EMPTY_HYPERLINK_SPAN

    unknown = session.remove_hyperlink("hl:body:deadbeef")
    assert not unknown.success
    assert unknown.error is not None
    assert unknown.error.code is EditErrorCode.HYPERLINK_NOT_FOUND

    absent = session.rename_bookmark("NeverExisted", "Whatever")
    assert not absent.success
    assert absent.error is not None
    assert absent.error.code is EditErrorCode.BOOKMARK_NOT_FOUND

    assert session.list_hyperlinks() == ()


def test_scoped_listing_sees_only_the_requested_story(
    session: DocxSession, paragraphs: list[str]
) -> None:
    assert session.add_hyperlink(
        paragraphs[0], CharSpan(0, 5), HyperlinkKind.EXTERNAL, "https://example.test/body"
    ).success
    assert session.set_header_text(paragraphs[0], HeaderFooterKind.DEFAULT, "header line").success
    header = next(
        a.id
        for a in session.project().anchor_index.values()
        if a.scope.startswith("hdr") and a.kind in ("p", "h", "li")
    )
    assert session.add_hyperlink(
        header, CharSpan(0, 6), HyperlinkKind.EXTERNAL, "https://example.test/header"
    ).success

    body_only = session.list_hyperlinks(ProjectionScopes.BODY)
    header_only = session.list_hyperlinks(ProjectionScopes.HEADERS)
    assert [link.scope for link in body_only] == ["body"]
    assert all(link.scope.startswith("hdr") for link in header_only)
    assert len(header_only) == 1
    assert len(session.list_hyperlinks()) == 2


def test_tracked_render_inline_rejects_metadata_mutations(
    session: DocxSession, paragraphs: list[str]
) -> None:
    session.set_tracked_changes(TrackedChangeMode.RENDER_INLINE)

    link = session.add_hyperlink(
        paragraphs[0], CharSpan(0, 5), HyperlinkKind.EXTERNAL, "https://example.test/x"
    )
    assert not link.success
    assert link.error is not None
    assert link.error.code is EditErrorCode.TRACKED_OPERATION_UNSUPPORTED

    bookmark = session.add_bookmark(
        "Tracked", DocumentRange(paragraphs[0], 0, paragraphs[0], 5)
    )
    assert not bookmark.success
    assert bookmark.error is not None
    assert bookmark.error.code is EditErrorCode.TRACKED_OPERATION_UNSUPPORTED

    assert session.list_hyperlinks() == ()
    assert all(b.name != "Tracked" for b in session.list_bookmarks())
