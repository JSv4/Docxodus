"""Issue #762 through the stdio host: the per-occurrence operation matrix, WebP writes, tight
wrap polygons, the explicit embed-linked conversion, and tracked image mutations.

The fixture is ``TestFiles/IM762-ImageCoverage.docx``: an embedded PNG in the first paragraph
and an external linked picture in the second.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from docx_scalpel import (
    DocxSession,
    DocxSessionSettings,
    EditErrorCode,
    FloatingImageLayout,
    ImageBinaryFormat,
    ImageInsertOptions,
    ImagePlacement,
    ImageWrapMode,
    ImageWrapPoint,
    ImageWrapPolygon,
    TrackedChangeMode,
    open_session,
)


@pytest.fixture(scope="session")
def coverage_bytes(test_files_dir: Path) -> bytes:
    return (test_files_dir / "IM762-ImageCoverage.docx").read_bytes()


def _png(width: int, height: int) -> bytes:
    return (b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR"
            + width.to_bytes(4, "big") + height.to_bytes(4, "big"))


def _webp(width: int, height: int) -> bytes:
    w, h = width - 1, height - 1
    header = bytes([w & 0xFF, ((w >> 8) & 0x3F) | ((h & 0x03) << 6), (h >> 2) & 0xFF, (h >> 10) & 0x0F])
    return b"RIFF" + (22).to_bytes(4, "little") + b"WEBP" + b"VP8L" + (10).to_bytes(4, "little") + b"\x2f" + header + b"\x00" * 5


def _linked(session: DocxSession):
    return next(image for image in session.list_images() if image.is_linked)


def _embedded(session: DocxSession):
    return next(image for image in session.list_images() if not image.is_linked)


def test_capabilities_and_matrix_are_typed(coverage_bytes: bytes) -> None:
    with open_session(coverage_bytes) as session:
        capabilities = session.get_image_capabilities()
        assert "embed_linked" in capabilities.operations
        assert ImageWrapMode.TIGHT in capabilities.mutable_wrap_modes
        assert "set_floating_layout" in capabilities.tracked_operations
        assert next(m for m in capabilities.markups if m.markup == "linked_picture").operations == (
            "embed_linked", "set_dimensions", "set_metadata", "set_floating_layout", "remove")
        webp = next(f for f in capabilities.formats if f.format is ImageBinaryFormat.WEBP)
        assert webp.can_insert and webp.can_replace

        linked = _linked(session)
        assert not linked.can_mutate
        assert linked.operation("replace") is not None and not linked.operation("replace").can_mutate
        assert "embed_linked" in (linked.operation("replace").reason or "")
        assert linked.operation("embed_linked").can_mutate
        embedded = _embedded(session)
        assert embedded.can_mutate
        assert embedded.operation("embed_linked").reason == "picture is already embedded"


def test_embed_linked_converts_the_picture_and_webp_is_writable(coverage_bytes: bytes) -> None:
    with open_session(coverage_bytes) as session:
        linked = _linked(session)
        refused = session.replace_image(linked.id, _png(6, 7))
        assert refused.error is not None and refused.error.code is EditErrorCode.LINKED_IMAGE_READ_ONLY
        result = session.embed_linked_image(linked.id, _webp(6, 7))
        assert result.success, result.error
        converted = next(image for image in session.list_images() if image.id == linked.id)
        assert not converted.is_linked and converted.is_embedded and converted.can_mutate
        assert converted.format is ImageBinaryFormat.WEBP
        assert converted.content_type == "image/webp"
        assert converted.intrinsic_width_pixels == 6
        assert session.undo()
        assert _linked(session).linked_target == "https://example.test/linked.png"


def test_tight_wrap_polygon_round_trips(coverage_bytes: bytes) -> None:
    with open_session(coverage_bytes) as session:
        anchor = _embedded(session).anchor_id
        outline = ImageWrapPolygon((ImageWrapPoint(0, 10800), ImageWrapPoint(10800, 0),
                                    ImageWrapPoint(21600, 10800), ImageWrapPoint(0, 10800)), edited=True)
        inserted = session.insert_image(anchor, 0, _png(4, 5), ImageInsertOptions(
            placement=ImagePlacement.FLOATING,
            floating_layout=FloatingImageLayout(wrap_mode=ImageWrapMode.TIGHT, wrap_polygon=outline)))
        assert inserted.success, inserted.error
        image = next(image for image in session.list_images() if image.id == inserted.image_id)
        assert image.floating_layout_supported
        assert image.floating_layout is not None
        assert image.floating_layout.wrap_mode is ImageWrapMode.TIGHT
        assert image.floating_layout.wrap_polygon == outline
        # Leaving the polygon out means the picture rectangle.
        through = session.set_image_floating_layout(image.id, FloatingImageLayout(wrap_mode=ImageWrapMode.THROUGH))
        assert through.success, through.error
        image = next(image for image in session.list_images() if image.id == inserted.image_id)
        assert image.floating_layout.wrap_polygon.points[2] == ImageWrapPoint(21600, 21600)
        assert not image.floating_layout.wrap_polygon.edited


def test_tracked_replace_records_a_revision_pair(coverage_bytes: bytes) -> None:
    settings = DocxSessionSettings(tracked_changes=TrackedChangeMode.RENDER_INLINE)
    with open_session(coverage_bytes, settings) as session:
        embedded = _embedded(session)
        assert embedded.can_mutate
        result = session.replace_image(embedded.id, _png(9, 9))
        assert result.success, result.error
        assert result.image_id != embedded.id
        images = {image.id: image for image in session.list_images()}
        assert not images[embedded.id].can_mutate
        assert "tracked deletion" in (images[embedded.id].unsupported_reason or "")
        assert images[result.image_id].can_mutate
        assert images[result.image_id].intrinsic_width_pixels == 9
        assert {"insert", "delete"} <= {revision.type for revision in session.list_revisions()}
        assert session.reject_all_revisions().success
        restored = _embedded(session)
        assert restored.id == embedded.id
        assert restored.intrinsic_width_pixels == 2
