"""Coverage for the block-metadata read surface on DocxSession.

Mirrors the BM00x .NET tests in spirit but exercises the Python wrapper end-to-end
through the stdio host. Uses the shared ``test_files_dir`` fixture from
``conftest.py`` to locate the byte-identical TestFiles corpus.
"""

from __future__ import annotations

from pathlib import Path
from typing import Iterator

import pytest

from docx_scalpel import DocxSession, FormatOp, open_session
from docx_scalpel.types import BlockMetadata, FormattingInspection, NumberFormat, StyleInfo


@pytest.fixture
def list_session(test_files_dir: Path) -> Iterator[DocxSession]:
    fixture = test_files_dir / "DB012-Lists-With-Different-Numberings.docx"
    if not fixture.exists():
        pytest.skip(f"fixture missing: {fixture}")
    session = open_session(fixture.read_bytes())
    try:
        yield session
    finally:
        session.close()


def _first_anchor_of_kind(session: DocxSession, kind: str):
    projection = session.project()
    for anchor in projection.anchor_index.values():
        if anchor.kind == kind:
            return anchor
    return None


def test_get_block_metadata_plain_paragraph(list_session: DocxSession) -> None:
    para = _first_anchor_of_kind(list_session, "p")
    if para is None:
        pytest.skip("fixture has no plain paragraph anchors")
    meta = list_session.get_block_metadata(para.id)
    assert isinstance(meta, BlockMetadata)
    assert meta.kind == "p"
    assert meta.scope == "body"


def test_get_block_metadata_unknown_anchor_returns_none(list_session: DocxSession) -> None:
    assert list_session.get_block_metadata("p:body:does-not-exist") is None


def test_get_block_metadatas_bulk_dedups(list_session: DocxSession) -> None:
    para = _first_anchor_of_kind(list_session, "p")
    if para is None:
        pytest.skip("fixture has no plain paragraph anchors")
    result = list_session.get_block_metadatas([para.id, para.id, "p:body:missing"])
    assert len(result) == 2
    assert result[para.id] is not None
    assert result["p:body:missing"] is None


def test_get_list_membership_li_anchor(list_session: DocxSession) -> None:
    li = _first_anchor_of_kind(list_session, "li")
    if li is None:
        pytest.skip("fixture has no list-item anchors")
    membership = list_session.get_list_membership(li.id)
    assert membership is not None
    assert membership.num_id > 0
    assert membership.level >= 0
    assert isinstance(membership.format, NumberFormat)
    assert membership.anchor_id == li.id
    assert membership.start >= 0


def test_get_section_info_body_anchor(list_session: DocxSession) -> None:
    para = _first_anchor_of_kind(list_session, "p")
    if para is None:
        para = _first_anchor_of_kind(list_session, "li")
    if para is None:
        pytest.skip("fixture has no body anchors")
    info = list_session.get_section_info(para.id)
    assert info is not None
    assert info.anchor_id == para.id
    assert info.page_width_twips > 0
    assert info.columns >= 1


def test_style_and_direct_effective_formatting_introspection(list_session: DocxSession) -> None:
    styles = list_session.list_styles()
    assert styles
    assert all(isinstance(style, StyleInfo) for style in styles)

    para = next(
        (
            anchor
            for anchor in list_session.project().anchor_index.values()
            if anchor.kind in ("p", "li") and anchor.text_preview
        ),
        None,
    )
    if para is None:
        pytest.skip("fixture has no paragraph-like anchors")
    formatting = list_session.get_formatting(para.id)
    assert isinstance(formatting, FormattingInspection)
    assert formatting.anchor_id == para.id
    # The effective branch fills `alignment` unconditionally, so asserting it is non-None
    # cannot fail. Assert something that discriminates instead: the resolved effective style
    # id must be a style the document actually declares (it is `pStyle` or the document's
    # default paragraph style, both of which ListStyles enumerates).
    effective_style_id = formatting.effective_paragraph.style_id
    if effective_style_id is not None:
        assert effective_style_id in {style.id for style in styles}
    # The effective layer resolves toggles and line spacing that the direct layer leaves
    # absent; if it ever stopped, this fails rather than silently reporting "inherit" as false.
    assert formatting.effective_paragraph.keep_next is not None
    assert formatting.effective_paragraph.line_spacing is not None

    spans = list_session.list_inline_spans(para.id)
    if not spans:
        pytest.skip("fixture paragraph has no text-bearing runs")
    first = spans[0]
    assert first.anchor_id == para.id
    result = list_session.apply_format(first.anchor_id, first.span, FormatOp(bold=True))
    assert result.success
