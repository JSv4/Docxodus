"""Tests for ``DocxSession.fill_placeholders`` — Python mirror of C# ``DS240``–``DS246``.

Verifies the multi-pass loop, dollar-prefix preservation, skip behavior, and
``max_passes`` validation. Exercises the client-side loop end-to-end against the
real host so any drift between Python and the C# semantics surfaces here.
"""

from __future__ import annotations

import pytest

from docx_scalpel import (
    BulkEditResult,
    ContextBoundary,
    FillOptions,
    PlaceholderKind,
    PlaceholderKinds,
    Position,
    TemplatePlaceholder,
    open_session,
)


def _seed_doc(tour_plan_bytes: bytes) -> bytes:
    """Inject a BlankFill / dollar-prefix / nested-bracket triplet into the HC001
    fixture so the fill loop has every placeholder shape to chew on.

    Builds a synthetic test corpus on top of an existing fixture instead of
    hand-rolling OOXML — the Python wrapper doesn't have a DocumentFormat.OpenXml
    equivalent, and round-tripping through the session keeps the doc shape valid.
    The inserts land at top-level body (above the first body anchor), not inside
    a table, so the placeholders end up in standalone paragraphs that
    ``find_placeholders`` and ``get_anchor_info`` both see directly.
    """
    with open_session(tour_plan_bytes) as session:
        proj = session.project()
        body_anchors = sorted(
            (t for t in proj.anchor_index.values()
             if t.kind in ("p", "h") and t.scope == "body"),
            key=lambda t: proj.markdown.find("{#" + t.id + "}"),
        )
        first = body_anchors[0]
        session.insert_paragraph(
            first.id, Position.BEFORE,
            "The name of this corporation is [_____].",
        )
        session.insert_paragraph(
            first.id, Position.BEFORE,
            "The price per share is $[___].",
        )
        session.insert_paragraph(
            first.id, Position.BEFORE,
            "[outer [inner] clause]",
        )
        return session.save()


def _placeholder_anchor_ids(session) -> set[str]:
    return {p.match.enclosing_anchor.id for p in session.find_placeholders()}


def test_fill_placeholders_blank_fill_replaces(tour_plan_bytes: bytes) -> None:
    seeded = _seed_doc(tour_plan_bytes)
    with open_session(seeded) as session:
        ids_before = _placeholder_anchor_ids(session)
        # Find the paragraph that hosts the BlankFill so we can verify its
        # text changed.
        blank_anchor = next(
            p.match.enclosing_anchor.id
            for p in session.find_placeholders()
            if p.kind == PlaceholderKind.BLANK_FILL and p.match.text == "[_____]"
        )

        result = session.fill_placeholders(
            lambda p: "ACME, INC." if (
                p.kind == PlaceholderKind.BLANK_FILL
                and p.match.text == "[_____]"
            ) else None,
        )
        assert isinstance(result, BulkEditResult)
        assert result.filled >= 1
        assert result.errors == ()

        info = session.get_anchor_info(blank_anchor)
        assert info is not None
        assert "ACME, INC." in info.text_preview
        assert "[_____]" not in info.text_preview

        # The BlankFill placeholder is gone; the other placeholders the picker
        # said `None` to remain available.
        remaining_blanks = [
            p for p in session.find_placeholders()
            if p.kind == PlaceholderKind.BLANK_FILL and p.match.text == "[_____]"
        ]
        assert remaining_blanks == []
        assert ids_before, "test setup: seeded doc should have had placeholders"


def test_fill_placeholders_picker_returning_none_skips(tour_plan_bytes: bytes) -> None:
    seeded = _seed_doc(tour_plan_bytes)
    with open_session(seeded) as session:
        result = session.fill_placeholders(lambda p: None)
        assert result.filled == 0
        assert result.skipped > 0
        assert len(result.unfilled) == result.skipped
        # Every unfilled entry is a TemplatePlaceholder, not a stringly-typed shape.
        assert all(isinstance(u, TemplatePlaceholder) for u in result.unfilled)
        assert result.passes == 0  # nothing was filled
        # Picker said no to everything → every placeholder is still present.
        # `still_present` should equal the post-loop count.
        assert result.still_present == result.skipped


def test_fill_placeholders_preserves_dollar_prefix(tour_plan_bytes: bytes) -> None:
    seeded = _seed_doc(tour_plan_bytes)
    with open_session(seeded) as session:
        dollar_anchor = next(
            p.match.enclosing_anchor.id
            for p in session.find_placeholders()
            if "$[___]" in p.match.text
        )

        result = session.fill_placeholders(
            lambda p: "0.20" if "$[___]" in p.match.text else None,
        )
        assert result.filled == 1

        info = session.get_anchor_info(dollar_anchor)
        assert info is not None
        assert "$0.20" in info.text_preview
        assert "$$0.20" not in info.text_preview  # dollar wasn't double-prepended


def test_fill_placeholders_dollar_prefix_disabled(tour_plan_bytes: bytes) -> None:
    seeded = _seed_doc(tour_plan_bytes)
    with open_session(seeded) as session:
        dollar_anchor = next(
            p.match.enclosing_anchor.id
            for p in session.find_placeholders()
            if "$[___]" in p.match.text
        )

        result = session.fill_placeholders(
            lambda p: "0.20" if "$[___]" in p.match.text else None,
            FillOptions(preserve_dollar_prefix=False),
        )
        assert result.filled == 1

        info = session.get_anchor_info(dollar_anchor)
        assert info is not None
        # With the option disabled, the captured "$" prefix in the span is
        # overwritten by the replacement.
        assert "share is 0.20." in info.text_preview
        assert "$0.20" not in info.text_preview


def test_fill_placeholders_alternative_clause_multipass(tour_plan_bytes: bytes) -> None:
    """Nested ``[outer [inner] clause]`` requires multiple passes to converge."""
    seeded = _seed_doc(tour_plan_bytes)

    def unwrap_brackets(p: TemplatePlaceholder) -> str | None:
        t = p.match.text
        lb, rb = t.find("["), t.rfind("]")
        if lb < 0 or rb <= lb:
            return None
        return t[:lb] + t[lb + 1:rb] + t[rb + 1:]

    with open_session(seeded) as session:
        result = session.fill_placeholders(
            unwrap_brackets,
            FillOptions(kinds=PlaceholderKinds.ALTERNATIVE_CLAUSE),
        )
        # The nested clause needs at least two passes to fully unwrap.
        assert result.passes >= 2
        # Single-call done-check: no AlternativeClause matches remain.
        assert result.still_present == 0
        leftover = session.find_placeholders(PlaceholderKinds.ALTERNATIVE_CLAUSE)
        assert leftover == ()


def test_fill_placeholders_max_passes_rejects_zero_or_negative(tour_plan_bytes: bytes) -> None:
    with open_session(tour_plan_bytes) as session:
        with pytest.raises(ValueError, match="max_passes"):
            session.fill_placeholders(lambda p: None, FillOptions(max_passes=0))
        with pytest.raises(ValueError, match="max_passes"):
            session.fill_placeholders(lambda p: None, FillOptions(max_passes=-1))


def test_fill_placeholders_passes_zero_when_no_matches(tour_plan_bytes: bytes) -> None:
    """``passes == 0`` is the well-defined signal for "nothing was filled"."""
    with open_session(tour_plan_bytes) as session:
        # HC001 has only AlternativeClause placeholders; ask only for BlankFill.
        result = session.fill_placeholders(
            lambda p: "X",
            FillOptions(kinds=PlaceholderKinds.BLANK_FILL),
        )
        assert result.filled == 0
        assert result.passes == 0


def test_fill_placeholders_default_kinds_visits_alternative_clause(tour_plan_bytes: bytes) -> None:
    """Default ``kinds=ALL`` invokes the picker for AlternativeClause too — the
    foot-gun the C# FillOptions.Kinds default change closed (DS244a)."""
    with open_session(tour_plan_bytes) as session:
        seen_kinds: set[PlaceholderKind] = set()

        def remember(p: TemplatePlaceholder) -> str | None:
            seen_kinds.add(p.kind)
            return None

        session.fill_placeholders(remember)
        # HC001's body uses AlternativeClause placeholders; without the default
        # widening to ALL, the picker would never have seen them.
        assert PlaceholderKind.ALTERNATIVE_CLAUSE in seen_kinds
