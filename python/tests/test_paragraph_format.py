"""End-to-end test for the paragraph-format surface through docx-scalpel (issue #301).

``set_paragraph_format`` was missing from ``docx-scalpel`` entirely — no ``Dispatcher.cs``
case, no ``ParagraphFormatOp``/``ParagraphBorderEdge`` types, no ``DocxSession`` method — so
alignment, indent, page-break-before, and paragraph borders (``w:pBdr``) were all unreachable
from Python even though the underlying `DocxSession.SetParagraphFormat` already supported them.
This proves the full op, including the border fields, round-trips into the OOXML.
"""

from __future__ import annotations

import zipfile
from io import BytesIO
from pathlib import Path

import pytest

from docx_scalpel import (
    ParagraphAlignment,
    ParagraphBorderEdge,
    ParagraphFormatOp,
    open_session,
)

FIXTURE = Path(__file__).parents[2] / "TestFiles" / "DA001-TemplateDocument.docx"


def _first_body_anchor(session) -> str:
    for anchor_id in session.project().anchor_index:
        if anchor_id.startswith("p:body:"):
            return anchor_id
    pytest.fail("no body paragraph in fixture")


def _document_xml(saved: bytes) -> str:
    with zipfile.ZipFile(BytesIO(saved)) as zf:
        return zf.read("word/document.xml").decode("utf-8")


def test_set_paragraph_format_alignment_indent_and_page_break():
    with open_session(FIXTURE.read_bytes()) as session:
        anchor = _first_body_anchor(session)

        result = session.set_paragraph_format(
            anchor,
            ParagraphFormatOp(
                alignment=ParagraphAlignment.CENTER,
                indent_delta=720,
                page_break_before=True,
            ),
        )
        assert result.success

        xml = _document_xml(session.save())
        assert 'w:jc w:val="both"' not in xml  # sanity: not silently coerced to justify
        assert 'w:jc w:val="center"' in xml
        assert "w:pageBreakBefore" in xml
        assert 'w:left="720"' in xml


def test_set_paragraph_format_adds_and_clears_borders():
    # The fixture already carries w:pBdr on OTHER paragraphs, so assertions target this one
    # paragraph's own XML (session.raw.get_xml) rather than scanning the whole document.
    with open_session(FIXTURE.read_bytes()) as session:
        anchor = _first_body_anchor(session)

        added = session.set_paragraph_format(
            anchor,
            ParagraphFormatOp(
                top_border=ParagraphBorderEdge(style="double", size=18, color="FF0000"),
                bottom_border=ParagraphBorderEdge(style="single", size=6),
            ),
        )
        assert added.success

        xml = session.raw.get_xml(anchor)
        assert "w:pBdr" in xml
        assert 'w:top w:val="double"' in xml
        assert 'w:sz="18"' in xml
        assert 'w:color="FF0000"' in xml
        assert 'w:bottom w:val="single"' in xml
        # bottom_border omitted style/size defaults: sz falls back to 6, color to "auto".
        assert 'w:sz="6"' in xml
        assert 'w:color="auto"' in xml

        cleared = session.set_paragraph_format(
            anchor, ParagraphFormatOp(clear_borders=True)
        )
        assert cleared.success

        cleared_xml = session.raw.get_xml(anchor)
        assert "w:pBdr" not in cleared_xml


def test_set_paragraph_format_unknown_anchor_fails():
    with open_session(FIXTURE.read_bytes()) as session:
        result = session.set_paragraph_format(
            "{#p:body:doesnotexist000000000000000000}",
            ParagraphFormatOp(alignment=ParagraphAlignment.LEFT),
        )
        assert not result.success
        assert result.error.code.value == "anchor_not_found"
