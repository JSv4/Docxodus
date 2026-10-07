"""DOCX->HTML conversion on the Python surface -- stateless + session-bound."""

from __future__ import annotations

from docx_scalpel import HtmlOptions, convert_docx_to_html, docx_diff_compare, open_session


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


def test_th004_word_revision_presentation_colours_authors(test_files_dir) -> None:
    # Issue #851: revision_presentation=1 draws tracked changes in Word's All Markup style.
    original = (test_files_dir / "CA" / "CA001-Plain.docx").read_bytes()
    modified = (test_files_dir / "CA" / "CA001-Plain-Mod.docx").read_bytes()
    redline = docx_diff_compare(original, modified)

    word = convert_docx_to_html(
        redline, HtmlOptions(render_tracked_changes=True, revision_presentation=1)
    )
    default = convert_docx_to_html(redline, HtmlOptions(render_tracked_changes=True))

    assert "rev-author-0" in word
    assert "rev-changed-line" in word
    assert "rev-author-" not in default


def test_th005_revision_presentation_on_the_wire() -> None:
    assert HtmlOptions().to_wire()["revisionPresentation"] == 0
    assert HtmlOptions(revision_presentation=1).to_wire()["revisionPresentation"] == 1


def test_th006_semantic_lists_emit_list_markup(test_files_dir) -> None:
    # Issue #895: semantic_lists turns Word list paragraphs into <ol>/<li>.
    numbered = (test_files_dir / "DB012-Lists-With-Different-Numberings.docx").read_bytes()

    semantic = convert_docx_to_html(numbered, HtmlOptions(semantic_lists=True))
    default = convert_docx_to_html(numbered)

    assert "<ol" in semantic and "<li" in semantic
    assert "<ol" not in default and "<li" not in default


def test_th007_semantic_lists_on_the_wire() -> None:
    assert HtmlOptions().to_wire()["semanticLists"] is False
    assert HtmlOptions(semantic_lists=True).to_wire()["semanticLists"] is True
