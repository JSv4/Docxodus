"""End-to-end tests for reference-field authoring (issue #607) through docx-scalpel.

The three tables narrowing the library to the DOCX toolchain took away with ``ReferenceAdder``:
contents, figures, authorities. The point of the ops is that a caller never writes a switch
string, so these assert the switch string — a malformed one renders as *nothing* in Word,
silently, and no schema check catches it.
"""

from __future__ import annotations

import re
import zipfile
from io import BytesIO
from pathlib import Path

import pytest

from docx_scalpel import AuthorityCategory, open_session
from docx_scalpel.enums import EditErrorCode

FIXTURE = Path(__file__).parents[2] / "TestFiles" / "DA001-TemplateDocument.docx"


def _first_body_anchor(session) -> str:
    for anchor_id in session.project().anchor_index:
        if anchor_id.startswith("p:body:"):
            return anchor_id
    pytest.fail("no body paragraph in fixture")


def _document_xml(saved: bytes) -> str:
    with zipfile.ZipFile(BytesIO(saved)) as zf:
        return zf.read("word/document.xml").decode("utf-8")


def _instructions(saved: bytes) -> list[str]:
    return [
        m.strip()
        for m in re.findall(r"<w:instrText[^>]*>(.*?)</w:instrText>", _document_xml(saved))
    ]


def test_table_of_contents_writes_a_dirty_toc_field_and_asks_word_to_update():
    with open_session(FIXTURE.read_bytes()) as session:
        anchor = _first_body_anchor(session)

        result = session.insert_table_of_contents(anchor)
        assert result.success, result.error

        saved = session.save()
        assert _instructions(saved) == ['TOC \\o "1-3" \\h \\z \\u']
        assert 'w:fldCharType="begin" w:dirty="true"' in _document_xml(saved)

        with zipfile.ZipFile(BytesIO(saved)) as zf:
            settings = zf.read("word/settings.xml").decode("utf-8")
        assert "updateFields" in settings


def test_typed_options_become_switches_without_the_caller_writing_one():
    with open_session(FIXTURE.read_bytes()) as session:
        anchor = _first_body_anchor(session)

        assert session.insert_table_of_contents(
            anchor, levels="1-2", hyperlinks=False, use_outline_levels=False, title=None
        ).success
        assert _instructions(session.save()) == ['TOC \\o "1-2" \\z']


def test_table_of_figures_and_authorities_carry_their_own_switches():
    with open_session(FIXTURE.read_bytes()) as session:
        anchor = _first_body_anchor(session)

        assert session.insert_table_of_figures(anchor, caption_label="Exhibit").success
        assert _instructions(session.save()) == ['TOC \\c "Exhibit" \\h']

    with open_session(FIXTURE.read_bytes()) as session:
        anchor = _first_body_anchor(session)

        assert session.insert_table_of_authorities(
            anchor, category=AuthorityCategory.STATUTES
        ).success
        # The wire name hides Word's number; the field carries the number.
        assert _instructions(session.save()) == ['TOA \\c "2" \\h']


def test_malformed_levels_are_refused_without_touching_the_document():
    with open_session(FIXTURE.read_bytes()) as session:
        anchor = _first_body_anchor(session)
        before = session.save()

        result = session.insert_table_of_contents(anchor, levels="0-3")

        assert not result.success
        assert result.error is not None
        assert result.error.code == EditErrorCode.INVALID_REFERENCE_FIELD
        assert session.save() == before
