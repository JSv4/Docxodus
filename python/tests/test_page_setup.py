"""End-to-end tests for the page-setup surface behind the editor's Word-parity work through
docx-scalpel: ``set_page_setup`` / ``SectionInfo`` read-back, the first-page / odd-even story
flags (``set_header_footer_kind_enabled``), the numeric comment id, and the highlight /
caps / small-caps ``FormatOp`` fields.
"""

from __future__ import annotations

import zipfile
from io import BytesIO
from pathlib import Path

import pytest

from docx_scalpel import FormatOp, PageSetupOp, open_session
from docx_scalpel.enums import EditErrorCode, HeaderFooterKind

FIXTURE = Path(__file__).parents[2] / "TestFiles" / "DA001-TemplateDocument.docx"


def _first_body_anchor(session) -> str:
    for anchor_id in session.project().anchor_index:
        if anchor_id.startswith("p:body:"):
            return anchor_id
    pytest.fail("no body paragraph in fixture")


def _part(saved: bytes, name: str) -> str:
    with zipfile.ZipFile(BytesIO(saved)) as zf:
        return zf.read(name).decode("utf-8")


def test_set_page_setup_margins_and_landscape_round_trip():
    with open_session(FIXTURE.read_bytes()) as session:
        anchor = _first_body_anchor(session)
        before = session.get_section_info(anchor)
        assert before is not None
        assert before.header_distance_twips > 0
        assert before.footer_distance_twips > 0

        result = session.set_page_setup(
            anchor,
            margin_top_twips=720,
            margin_bottom_twips=720,
            header_distance_twips=360,
            footer_distance_twips=360,
        )
        assert result.success, result.error

        after = session.get_section_info(anchor)
        assert after.margin_top_twips == 720
        assert after.margin_bottom_twips == 720
        assert after.header_distance_twips == 360
        assert after.footer_distance_twips == 360
        # Untouched attributes keep their values.
        assert after.margin_left_twips == before.margin_left_twips

        document_xml = _part(session.save(), "word/document.xml")
        assert 'w:header="360"' in document_xml
        assert 'w:footer="360"' in document_xml

        # A typed op works the same way, and landscape swaps a portrait sheet.
        assert session.set_page_setup(anchor, PageSetupOp(landscape=True)).success
        rotated = session.get_section_info(anchor)
        assert rotated.landscape is True
        assert rotated.page_width_twips == before.page_height_twips
        assert rotated.page_height_twips == before.page_width_twips
        assert 'w:orient="landscape"' in _part(session.save(), "word/document.xml")


def test_set_page_setup_rejects_impossible_geometry():
    with open_session(FIXTURE.read_bytes()) as session:
        anchor = _first_body_anchor(session)
        saved_before = _part(session.save(), "word/document.xml")

        too_wide = session.set_page_setup(anchor, margin_left_twips=9000, margin_right_twips=9000)
        assert not too_wide.success
        assert too_wide.error is not None
        assert too_wide.error.code is EditErrorCode.INVALID_PAGE_SETUP

        negative = session.set_page_setup(anchor, margin_top_twips=-1)
        assert negative.error.code is EditErrorCode.INVALID_PAGE_SETUP

        assert _part(session.save(), "word/document.xml") == saved_before


def test_header_footer_kind_flags_round_trip():
    with open_session(FIXTURE.read_bytes()) as session:
        anchor = _first_body_anchor(session)
        assert session.set_header_text(anchor, HeaderFooterKind.FIRST, "First page").success
        assert session.get_section_info(anchor).title_page is True

        off = session.set_header_footer_kind_enabled(anchor, HeaderFooterKind.FIRST, False)
        assert off.success, off.error
        info = session.get_section_info(anchor)
        assert info.title_page is False
        # The story part survives the flag being cleared, as in Word.
        assert any(ref.kind is HeaderFooterKind.FIRST for ref in info.header_refs)
        assert "w:titlePg" not in _part(session.save(), "word/document.xml")

        assert session.set_header_footer_kind_enabled(anchor, HeaderFooterKind.EVEN, True).success
        assert session.get_section_info(anchor).even_and_odd_headers is True
        assert "w:evenAndOddHeaders" in _part(session.save(), "word/settings.xml")
        assert session.set_header_footer_kind_enabled(anchor, HeaderFooterKind.EVEN, False).success
        assert session.get_section_info(anchor).even_and_odd_headers is False

        default = session.set_header_footer_kind_enabled(anchor, HeaderFooterKind.DEFAULT, False)
        assert not default.success
        assert default.error.code is EditErrorCode.INVALID_PAGE_SETUP


def test_list_comments_reports_numeric_id():
    with open_session(FIXTURE.read_bytes()) as session:
        anchor = _first_body_anchor(session)
        assert session.add_comment(anchor, None, "Reviewer", "Check this").success
        entry = session.list_comments()[0]
        assert entry.id >= 0
        assert f'w:id="{entry.id}"' in _part(session.save(), "word/comments.xml")


def test_format_op_highlight_caps_small_caps():
    with open_session(FIXTURE.read_bytes()) as session:
        anchor = _first_body_anchor(session)
        assert session.apply_format(anchor, None, FormatOp(highlight="yellow", caps=True)).success
        document_xml = _part(session.save(), "word/document.xml")
        assert 'w:highlight w:val="yellow"' in document_xml
        assert "<w:caps" in document_xml

        # Small caps evicts caps; "" clears the highlight.
        assert session.apply_format(anchor, None, FormatOp(highlight="", small_caps=True)).success
        document_xml = _part(session.save(), "word/document.xml")
        assert "w:highlight" not in document_xml
        assert "<w:caps" not in document_xml
        assert "<w:smallCaps" in document_xml

        bad = session.apply_format(anchor, None, FormatOp(highlight="pink"))
        assert not bad.success
