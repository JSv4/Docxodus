"""DOCX->HTML conversion on the Python surface -- stateless + session-bound."""

from __future__ import annotations

import re
from pathlib import Path

from docx_scalpel import HtmlOptions, convert_docx_to_html, open_session


def _document_order(markdown: str, anchor_id: str) -> int:
    """Position of an anchor's marker in the projection, for document-order sort.

    Mirrors the helper in ``test_smoke.py``; kept local for the same reason
    (conftest functions are not reliably importable by name).
    """
    i = markdown.find("{#" + anchor_id + "}")
    return i if i >= 0 else 1 << 30


def test_th001_stateless_convert_produces_html(tour_plan_bytes: bytes) -> None:
    html = convert_docx_to_html(tour_plan_bytes)
    assert "<html" in html
    assert "</html>" in html


def test_th002_css_prefix_option_applied(tour_plan_bytes: bytes) -> None:
    html = convert_docx_to_html(tour_plan_bytes, HtmlOptions(css_class_prefix="zz-"))
    assert "zz-" in html


def test_th003_session_to_html_reflects_edit(tour_plan_bytes: bytes) -> None:
    marker = "TH003UNIQUEMARKER"
    with open_session(tour_plan_bytes) as session:
        projection = session.project()
        # First body paragraph/heading/list-item anchor in document order.
        candidates = [
            t
            for t in projection.anchor_index.values()
            if t.kind in ("p", "h", "li") and t.scope == "body"
        ]
        anchor = min(
            candidates,
            key=lambda t: _document_order(projection.markdown, t.id),
        )
        result = session.replace_text(anchor.id, f"{marker} edited body.")
        assert result.success, result.error

        edited_html = session.to_html()
        assert marker in edited_html

    # Stateless conversion of the ORIGINAL bytes must not contain the edit.
    original_html = convert_docx_to_html(tour_plan_bytes)
    assert marker not in original_html


def test_th004_semantic_lists_is_off_unless_asked() -> None:
    assert HtmlOptions().to_wire()["semanticLists"] is False
    assert HtmlOptions(semantic_lists=True).to_wire()["semanticLists"] is True


def test_th005_semantic_lists_render_list_items(test_files_dir: Path) -> None:
    # Nine list paragraphs across three w:num instances.
    docx = (test_files_dir / "DB012-Lists-With-Different-Numberings.docx").read_bytes()
    item = re.compile(r"<li[\s>]")

    assert item.search(convert_docx_to_html(docx)) is None
    html = convert_docx_to_html(docx, HtmlOptions(semantic_lists=True))
    assert len(item.findall(html)) == 9
    assert len(re.findall(r"<ol[\s>]", html)) == 3
