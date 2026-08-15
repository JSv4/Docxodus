"""Native content controls end-to-end through the stdio host (issue #452).

`test_content_control_types.py` covers wire *decoding* only. This module drives every
one of the nine content-control routes across the real `docxodus-pyhost` subprocess, so
a typo in a route name or an argument key fails here instead of shipping: an unknown op
raises out of the host rather than returning an `EditResult`, and a misspelled argument
key raises a `FormatException` rather than reaching `DocxSessionOps`.

The fixture is HC030 — Word-authored, five controls (rich text, plain text, picture,
checkbox, combo box). It carries no date or repeating-section control, so those three
routes are proved reachable by asserting the *engine's* typed rejection of a
well-formed call against a wrong-typed target.
"""

from __future__ import annotations

from pathlib import Path
from typing import Iterator

import pytest

from docx_scalpel import (
    ContentControlBindingPolicy,
    ContentControlFillOptions,
    ContentControlPlacement,
    ContentControlType,
    DocxSession,
    ProjectionScopes,
    open_session,
)
from docx_scalpel.enums import EditErrorCode


def _png(width: int, height: int) -> bytes:
    """A minimal PNG signature + IHDR — enough for the host's format/dimension sniffing."""
    return (
        bytes([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A])
        + bytes([0, 0, 0, 13])
        + b"IHDR"
        + width.to_bytes(4, "big")
        + height.to_bytes(4, "big")
    )


@pytest.fixture(scope="session")
def content_control_bytes(test_files_dir: Path) -> bytes:
    return (test_files_dir / "HC030-Content-Controls.docx").read_bytes()


@pytest.fixture
def session(content_control_bytes: bytes) -> Iterator[DocxSession]:
    s = open_session(content_control_bytes)
    try:
        yield s
    finally:
        s.close()


def _only(session: DocxSession, kind: ContentControlType) -> str:
    matches = [c for c in session.list_content_controls() if c.type is kind]
    assert len(matches) == 1, f"expected exactly one {kind}, got {len(matches)}"
    return matches[0].anchor_id


def test_list_content_controls_decodes_the_word_authored_registry(session: DocxSession) -> None:
    controls = session.list_content_controls()

    assert [c.type for c in controls] == [
        ContentControlType.RICH_TEXT,
        ContentControlType.PLAIN_TEXT,
        ContentControlType.PICTURE,
        ContentControlType.CHECKBOX,
        ContentControlType.COMBO_BOX,
    ]
    assert all(c.anchor_id.startswith("sdt:body:") for c in controls)
    assert all(c.scope == "body" for c in controls)
    assert all(c.owning_part_uri.endswith("document.xml") for c in controls)
    assert [c.placement for c in controls] == [
        ContentControlPlacement.BLOCK,
        ContentControlPlacement.INLINE,
        ContentControlPlacement.BLOCK,
        ContentControlPlacement.BLOCK,
        ContentControlPlacement.BLOCK,
    ]
    assert all(c.has_valid_native_id and not c.has_duplicate_native_id for c in controls)
    assert all(c.can_mutate for c in controls), [c.unsupported_reason for c in controls]
    assert all(not c.is_bound and c.binding is None for c in controls)
    assert [c.item_values for c in controls if c.type is ContentControlType.COMBO_BOX] == [
        ("One", "Two", "Three")
    ]


def test_list_content_controls_honors_the_scopes_argument(session: DocxSession) -> None:
    assert len(session.list_content_controls(ProjectionScopes.BODY)) == 5
    assert session.list_content_controls(ProjectionScopes.HEADERS) == ()
    assert session.list_content_controls(ProjectionScopes.FOOTERS) == ()


def test_fill_text_and_rich_text_round_trip_through_the_host(session: DocxSession) -> None:
    plain = _only(session, ContentControlType.PLAIN_TEXT)
    rich = _only(session, ContentControlType.RICH_TEXT)

    text = session.fill_content_control_text(plain, "wired plain value")
    assert text.success, text.error
    assert [a.id for a in text.modified] == [plain]

    markdown = session.fill_content_control_rich_text(rich, "wired **rich** value")
    assert markdown.success, markdown.error

    by_anchor = {c.anchor_id: c for c in session.list_content_controls()}
    assert by_anchor[plain].text == "wired plain value"
    assert by_anchor[rich].text == "wired rich value"


def test_set_checked_and_select_item_persist_native_state(session: DocxSession) -> None:
    checkbox = _only(session, ContentControlType.CHECKBOX)
    combo = _only(session, ContentControlType.COMBO_BOX)

    checked = session.set_content_control_checked(checkbox, True)
    assert checked.success, checked.error

    selected = session.select_content_control_item(combo, "wired combo value")
    assert selected.success, selected.error

    by_anchor = {c.anchor_id: c for c in session.list_content_controls()}
    assert by_anchor[checkbox].text == "☒"
    assert by_anchor[combo].text == "wired combo value"


def test_fill_picture_accepts_base64_bytes(session: DocxSession) -> None:
    picture = _only(session, ContentControlType.PICTURE)

    result = session.fill_content_control_picture(picture, _png(4, 5))

    assert result.success, result.error
    assert [a.id for a in result.modified] == [picture]


def test_fill_picture_rejects_non_image_bytes_without_touching_the_session(
    session: DocxSession,
) -> None:
    picture = _only(session, ContentControlType.PICTURE)
    before = session.list_content_controls()

    result = session.fill_content_control_picture(picture, b"not an image at all")

    assert not result.success
    assert result.error is not None
    assert session.list_content_controls() == before


def test_set_date_route_is_wired_and_validates_its_value_argument(session: DocxSession) -> None:
    plain = _only(session, ContentControlType.PLAIN_TEXT)

    # A well-formed call reaches the engine, which rejects the wrong family.
    wrong_type = session.set_content_control_date(plain, "2026-08-14T00:00:00Z")
    assert wrong_type.error is not None
    assert wrong_type.error.code is EditErrorCode.CONTENT_CONTROL_WRONG_TYPE

    # displayText is optional and the value is parsed host-side, not passed through raw.
    bad_value = session.set_content_control_date(plain, "not-a-timestamp", "August 2026")
    assert bad_value.error is not None
    assert bad_value.error.code is EditErrorCode.INVALID_CONTENT_CONTROL_VALUE


def test_repeating_section_routes_are_wired_and_typed(session: DocxSession) -> None:
    plain = _only(session, ContentControlType.PLAIN_TEXT)

    add = session.add_repeating_section_item(plain)
    assert add.error is not None
    assert add.error.code is EditErrorCode.CONTENT_CONTROL_WRONG_TYPE

    add_after = session.add_repeating_section_item(plain, after_item_anchor_id=plain)
    assert add_after.error is not None
    assert add_after.error.code is EditErrorCode.CONTENT_CONTROL_WRONG_TYPE

    remove = session.remove_repeating_section_item(plain)
    assert remove.error is not None
    assert remove.error.code is EditErrorCode.CONTENT_CONTROL_WRONG_TYPE

    missing = session.remove_repeating_section_item("sdt:body:deadbeef")
    assert missing.error is not None
    assert missing.error.code is EditErrorCode.CONTENT_CONTROL_NOT_FOUND


def test_binding_policy_option_crosses_the_wire_on_every_fill(session: DocxSession) -> None:
    plain = _only(session, ContentControlType.PLAIN_TEXT)
    detach = ContentControlFillOptions(ContentControlBindingPolicy.DETACH_TARGET)

    # HC030 has no bindings, so detach_target is a no-op the fill must still accept.
    result = session.fill_content_control_text(plain, "detach-policy value", detach)

    assert result.success, result.error
    assert {c.anchor_id: c.text for c in session.list_content_controls()}[plain] == (
        "detach-policy value"
    )


def test_content_control_fills_are_undoable_and_survive_save_reopen(
    session: DocxSession, content_control_bytes: bytes
) -> None:
    plain = _only(session, ContentControlType.PLAIN_TEXT)
    original = {c.anchor_id: c.text for c in session.list_content_controls()}[plain]

    assert session.fill_content_control_text(plain, "persisted value").success
    saved = session.save()
    assert session.undo()
    assert {c.anchor_id: c.text for c in session.list_content_controls()}[plain] == original

    reopened = open_session(saved)
    try:
        # The anchor is derived from the native w:sdtPr/w:id, so it survives a clean save.
        assert {c.anchor_id: c.text for c in reopened.list_content_controls()}[plain] == (
            "persisted value"
        )
    finally:
        reopened.close()
