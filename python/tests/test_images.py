"""Native image CRUD and typed capability projection through the stdio host."""

from __future__ import annotations

from typing import Iterator

import pytest

from docx_scalpel import (
    DocxSession,
    ImageBinaryFormat,
    ImageDimensions,
    ImageInsertOptions,
    ImageMarkupKind,
    ProjectionScopes,
    open_session,
)


@pytest.fixture
def session(tour_plan_bytes: bytes) -> Iterator[DocxSession]:
    value = open_session(tour_plan_bytes)
    try:
        yield value
    finally:
        value.close()


def _first_body_paragraph(session: DocxSession) -> str:
    return next(
        anchor.id
        for anchor in session.project().anchor_index.values()
        if anchor.scope == "body" and anchor.kind in ("p", "h", "li")
    )


def _png_header(width: int, height: int) -> bytes:
    return (
        b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR"
        + width.to_bytes(4, "big")
        + height.to_bytes(4, "big")
    )


def test_capabilities_and_image_crud_are_typed(session: DocxSession) -> None:
    capabilities = session.get_image_capabilities()
    assert capabilities.default_dpi == 96
    assert capabilities.supports_network_fetch is False
    assert ImageBinaryFormat.PNG in {entry.format for entry in capabilities.formats}

    made = session.insert_image(
        _first_body_paragraph(session),
        0,
        _png_header(2, 3),
        ImageInsertOptions(width_points=72, alt_text="diagram"),
    )
    assert made.success, made.error
    image = session.list_images(ProjectionScopes.BODY)[0]
    assert image.id == made.image_id
    assert image.markup_kind is ImageMarkupKind.MODERN_DRAWING
    assert image.format is ImageBinaryFormat.PNG
    assert image.intrinsic_width_pixels == 2
    assert image.intrinsic_height_pixels == 3
    assert image.rendered_width_points == 72
    assert image.rendered_height_points == 108

    resized = session.set_image_dimensions(image.id, ImageDimensions(width_points=36))
    assert resized.success, resized.error
    assert session.list_images()[0].rendered_height_points == 54
    assert session.set_image_metadata(image.id, "updated", None).success

    reopened = open_session(session.save())
    try:
        persisted = reopened.list_images()[0]
        assert persisted.alt_text == "updated"
        assert reopened.remove_image(persisted.id).success
        assert reopened.list_images() == ()
    finally:
        reopened.close()
