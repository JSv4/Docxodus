"""A list's levels (Word's Define New Multilevel List) and Continue Numbering (#1790)."""

from __future__ import annotations

import pytest

from pathlib import Path

from docx_scalpel import ListFormat, ListLevelSpec, Position, open_session


class _Doc:
    """Paragraphs found by their text; an anchor's kind changes (p -> li) as a paragraph becomes
    a list item, so each is looked up again by the id's stable last part."""

    def __init__(self, *texts: str) -> None:
        fixture = Path(__file__).parents[2] / "TestFiles" / "DA001-TemplateDocument.docx"
        if not fixture.exists():
            pytest.skip(f"fixture missing: {fixture}")
        self.session = open_session(fixture.read_bytes())
        first = next(a for a in self.session.project().anchor_index.values() if a.scope == "body")
        result = self.session.insert_paragraph(first.id, Position.AFTER, "\n\n".join(texts))
        assert result.success, result.error
        assert len(result.created) == len(texts)
        self._unids = {text: anchor.id.rsplit(":", 1)[-1] for text, anchor in zip(texts, result.created)}

    def __getitem__(self, text: str) -> str:
        unid = self._unids[text]
        return next(a for a in self.session.project().anchor_index if a.endswith(":" + unid))

    def label(self, text: str) -> str | None:
        membership = self.session.get_list_membership(self[text])
        return membership.generated_label if membership else None


def test_get_and_set_list_levels() -> None:
    doc = _Doc("LA", "LB")
    session = doc.session
    try:
        assert session.apply_list_format_range(doc["LA"], doc["LB"], ListFormat.DECIMAL).success
        levels = session.get_list_levels(doc["LA"])
        assert levels is not None
        assert (levels[0].level, levels[0].format, levels[0].text) == (0, "decimal", "%1.")

        result = session.set_list_levels(
            doc["LA"], [ListLevelSpec(level=0, format="upperRoman", text="%1)", start=3)]
        )

        assert result.success, result.error
        level0 = session.get_list_levels(doc["LA"])[0]
        assert (level0.format, level0.text, level0.start) == ("upperRoman", "%1)", 3)
        assert doc.label("LA") == "III)"
        bad = session.set_list_levels(doc["LA"], [ListLevelSpec(level=1, text="%3.")])
        assert not bad.success
        assert bad.error.code.value == "invalid_list_level"
    finally:
        session.close()


def test_continue_numbering_joins_the_matching_list_only() -> None:
    doc = _Doc("A1", "A2", "Bullet", "N1", "N2")
    session = doc.session
    try:
        assert session.apply_list_format_range(doc["A1"], doc["A2"], ListFormat.DECIMAL).success
        assert session.apply_list_format(doc["Bullet"], ListFormat.BULLET).success
        assert session.apply_list_format_range(doc["N1"], doc["N2"], ListFormat.DECIMAL).success
        assert session.set_list_start_override(doc["N1"], 1).success
        assert doc.label("N1") == "1."

        result = session.continue_numbering(doc["N1"])

        assert result.success, result.error
        assert (doc.label("N1"), doc.label("N2")) == ("3.", "4.")
        refused = session.continue_numbering(doc["A1"])
        assert not refused.success
    finally:
        session.close()
