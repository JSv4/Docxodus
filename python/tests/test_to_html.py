"""DOCX->HTML conversion on the Python surface -- stateless + session-bound."""

from __future__ import annotations

import xml.etree.ElementTree as ET

from docx_scalpel import HtmlOptions, PageNumberField, convert_docx_to_html, docx_diff_compare, open_session
from docx_scalpel.enums import HeaderFooterKind


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


def test_th008_revision_visibility_wire_defaults() -> None:
    defaults = HtmlOptions().to_wire()
    assert defaults["showFormattingChanges"] is True
    assert defaults["showInsertionsAndDeletions"] is True
    hidden = HtmlOptions(show_formatting_changes=False, show_insertions_and_deletions=False).to_wire()
    assert hidden["showFormattingChanges"] is False
    assert hidden["showInsertionsAndDeletions"] is False


def test_th009_hidden_text_revisions_cross_the_stdio_boundary(test_files_dir) -> None:
    original = (test_files_dir / "CA" / "CA001-Plain.docx").read_bytes()
    modified = (test_files_dir / "CA" / "CA001-Plain-Mod.docx").read_bytes()
    redline = docx_diff_compare(original, modified)
    shown = ET.fromstring(convert_docx_to_html(redline, HtmlOptions(render_tracked_changes=True)))
    hidden = ET.fromstring(convert_docx_to_html(redline, HtmlOptions(
        render_tracked_changes=True, show_insertions_and_deletions=False)))
    tags = {"{http://www.w3.org/1999/xhtml}ins", "{http://www.w3.org/1999/xhtml}del"}
    assert any(node.tag in tags for node in shown.iter())
    assert not any(node.tag in tags for node in hidden.iter())


def test_th010_existing_positional_html_options_keep_their_meaning() -> None:
    options = HtmlOptions("Legacy", "zz-", False, ".legacy{}", 1, "c-", 1, .75, "p-",
                          True, 2, "a-", True, True, True, False, False, 1, True, "fr", True)
    wire = options.to_wire()
    assert wire["showDeletedContent"] is False
    assert wire["renderMoveOperations"] is False
    assert wire["revisionPresentation"] == 1
    assert wire["documentLanguage"] == "fr"
    assert wire["semanticLists"] is True
    assert wire["showFormattingChanges"] is True
    assert wire["showInsertionsAndDeletions"] is True


def test_th011_flow_page_fields_cross_stateless_and_session_apis(tour_plan_bytes) -> None:
    with open_session(tour_plan_bytes) as session:
        anchor = next(a for a in session.project().anchor_index if a.startswith("p:body:"))
        footer = session.set_footer_text(anchor, HeaderFooterKind.DEFAULT, "Page ")
        assert footer.success
        assert session.insert_page_number_field(footer.created[0].id, PageNumberField.CURRENT_PAGE).success
        saved = session.save()
        for convert in (session.to_html, lambda options: convert_docx_to_html(saved, options)):
            for stamp in (False, True):
                html = ET.fromstring(convert(HtmlOptions(
                    render_headers_and_footers=True, stamp_page_number_fields=stamp)))
                fields = [node.attrib["data-field"] for node in html.iter() if "data-field" in node.attrib]
                assert fields == (["PAGE"] if stamp else [])
