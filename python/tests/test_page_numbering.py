"""End-to-end tests for the page-numbering surface (issue #277) through docx-scalpel.

Two independent layers: the SECTION's ``w:pgNumType`` (``set_page_numbering`` /
``clear_page_numbering``, read back on ``SectionInfo``) and a per-field ``\\*``
general-formatting switch on ``insert_page_number_field``.
"""

from __future__ import annotations

import zipfile
from io import BytesIO
from pathlib import Path

import pytest

from docx_scalpel import NumberFormat, PageNumberField, open_session
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


def test_set_read_back_and_clear_page_numbering():
    with open_session(FIXTURE.read_bytes()) as session:
        anchor = _first_body_anchor(session)

        # Absent attributes read back as None, not as a fabricated decimal/1 default.
        before = session.get_section_info(anchor)
        assert before is not None
        assert before.page_number_start is None
        assert before.page_number_format is None

        assert session.set_page_numbering(
            anchor, start=1, format=NumberFormat.LOWER_ROMAN
        ).success

        after = session.get_section_info(anchor)
        assert after.page_number_start == 1
        assert after.page_number_format is NumberFormat.LOWER_ROMAN

        document_xml = _part(session.save(), "word/document.xml")
        assert 'w:start="1"' in document_xml
        assert 'w:fmt="lowerRoman"' in document_xml

        # Omitting a field leaves that attribute alone.
        assert session.set_page_numbering(anchor, format=NumberFormat.UPPER_ROMAN).success
        merged = session.get_section_info(anchor)
        assert merged.page_number_start == 1
        assert merged.page_number_format is NumberFormat.UPPER_ROMAN

        assert session.clear_page_numbering(anchor).success
        cleared = session.get_section_info(anchor)
        assert cleared.page_number_start is None
        assert cleared.page_number_format is None


def test_page_number_field_format_switch():
    with open_session(FIXTURE.read_bytes()) as session:
        anchor = _first_body_anchor(session)
        created = session.set_footer_text(anchor, HeaderFooterKind.DEFAULT, "Page ")
        assert created.success
        footer_anchor = created.created[0].id

        assert session.insert_page_number_field(
            footer_anchor, PageNumberField.CURRENT_PAGE, NumberFormat.LOWER_ROMAN
        ).success

        footer_xml = _part(session.save(), "word/footer1.xml")
        assert "PAGE \\* roman" in footer_xml
        # The cached result agrees with the switch — it is what a non-recomputing renderer shows.
        assert "<w:t>i</w:t>" in footer_xml


def test_page_numbering_rejects_non_page_values():
    with open_session(FIXTURE.read_bytes()) as session:
        anchor = _first_body_anchor(session)

        bullet = session.set_page_numbering(anchor, format=NumberFormat.BULLET)
        assert not bullet.success
        assert bullet.error.code is EditErrorCode.INVALID_PAGE_NUMBERING

        negative = session.set_page_numbering(anchor, start=-1)
        assert not negative.success
        assert negative.error.code is EditErrorCode.INVALID_PAGE_NUMBERING

        # Nothing was written by either rejected call.
        info = session.get_section_info(anchor)
        assert info.page_number_start is None
        assert info.page_number_format is None
