"""DOCX->HTML conversion on the Python surface -- stateless + session-bound."""

from __future__ import annotations

from docx_scalpel import HtmlOptions, convert_docx_to_html, open_session


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
            key=lambda t: (
                projection.markdown.find("{#" + t.id + "}")
                if projection.markdown.find("{#" + t.id + "}") >= 0
                else 1 << 30
            ),
        )
        result = session.replace_text(anchor.id, f"{marker} edited body.")
        assert result.success, result.error

        edited_html = session.to_html()
        assert marker in edited_html

    # Stateless conversion of the ORIGINAL bytes must not contain the edit.
    original_html = convert_docx_to_html(tour_plan_bytes)
    assert marker not in original_html
