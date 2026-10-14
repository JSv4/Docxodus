"""Coverage for the find-by surface on DocxSession.

Python mirror of ``npm/tests/find-by.spec.ts`` (Issue #171): exercises
``exists`` / ``find_by_text`` / ``find_all_by_text`` / ``find_by_regex`` /
``find_by_kind`` / ``replace_match`` end-to-end through the stdio host, so the
Python wrapper and the npm/TS wrapper are verified against the same shared
``DocxSessionOps`` core with byte-identical fixtures.

Fixture: HC006-Test-01.docx — English body text (so the regex needles have
Latin words to hit), multiple headings, and paragraphs. (HC001's body is CJK.)
"""

from __future__ import annotations

from pathlib import Path
from typing import Iterator

import pytest

from docx_scalpel import DocxSession, open_session
from docx_scalpel.types import AnchorTarget, FindOptions


@pytest.fixture
def english_session(test_files_dir: Path) -> Iterator[DocxSession]:
    fixture = test_files_dir / "HC006-Test-01.docx"
    if not fixture.exists():
        pytest.skip(f"fixture missing: {fixture}")
    session = open_session(fixture.read_bytes())
    try:
        yield session
    finally:
        session.close()


def _first_word(session: DocxSession) -> str:
    """A real ≥5-letter Latin word from the body, via grep — keeps needles out
    of hard-coded fixture text."""
    matches = session.grep(r"[A-Za-z]{5,}")
    if not matches:
        pytest.skip("fixture has no ≥5-letter words to search for")
    return matches[0].text


def test_exists_true_for_real_anchor_false_for_bogus(english_session: DocxSession) -> None:
    projection = english_session.project()
    real_id = next(iter(projection.anchor_index))
    assert english_session.exists(real_id) is True
    assert english_session.exists("p:body:deadbeefdeadbeef") is False


def test_find_by_text_locates_a_real_needle(english_session: DocxSession) -> None:
    needle = _first_word(english_session)

    single = english_session.find_by_text(needle)
    assert isinstance(single, AnchorTarget)
    assert single.id

    all_hits = english_session.find_all_by_text(needle)
    assert len(all_hits) >= 1
    assert any(h.id == single.id for h in all_hits)

    assert english_session.find_by_text("zzqx_no_such_needle_42") is None


def test_find_all_by_text_honors_ignore_case(english_session: DocxSession) -> None:
    needle = _first_word(english_session)
    upper = needle.upper()

    strict = english_session.find_all_by_text(upper)
    loose = english_session.find_all_by_text(upper, FindOptions(ignore_case=True))

    # Case-insensitive search never finds fewer anchors than the case-sensitive
    # one, and finds the original-cased occurrence(s).
    assert len(loose) >= len(strict)
    assert len(loose) >= 1


def test_find_by_text_kind_filter_restricts_results(english_session: DocxSession) -> None:
    # Derive the needle from a real heading so the kind filter has something to
    # discriminate on (avoids a vacuously-true assertion over an empty result).
    headings = english_session.find_by_kind("h", "body")
    if not headings:
        pytest.skip("fixture has no headings")
    heading = headings[0]
    word = max(heading.text_preview.split(), key=len, default="")
    if len(word) < 4:
        pytest.skip("heading has no usable word to search for")

    as_headings = english_session.find_all_by_text(
        word, FindOptions(ignore_case=True, kind_filter="h")
    )
    as_paragraphs = english_session.find_all_by_text(
        word, FindOptions(ignore_case=True, kind_filter="p")
    )

    # The heading the needle came from is found under the "h" filter...
    assert any(h.id == heading.id for h in as_headings)
    assert all(h.kind == "h" for h in as_headings)
    # ...and never leaks into the "p" filter.
    assert all(p.kind == "p" for p in as_paragraphs)
    assert heading.id not in {p.id for p in as_paragraphs}


def test_find_by_regex_returns_multiple_anchors(english_session: DocxSession) -> None:
    hits = english_session.find_by_regex(r"\S+")  # any non-whitespace text
    assert len(hits) >= 2
    ids = [h.id for h in hits]
    assert len(set(ids)) == len(ids)  # anchors are unique


def test_find_by_kind_paragraphs_and_heading_scope(english_session: DocxSession) -> None:
    paras = english_session.find_by_kind("p")
    assert len(paras) >= 1
    assert all(p.kind == "p" for p in paras)

    headings = english_session.find_by_kind("h", "body")
    assert len(headings) >= 1
    assert all(h.kind == "h" and h.scope == "body" for h in headings)


def test_replace_match_workflow(english_session: DocxSession) -> None:
    matches = english_session.grep(r"[A-Za-z]{5,}")
    if not matches:
        pytest.skip("fixture has no ≥5-letter words to replace")
    match = matches[0]
    original = match.text

    edit = english_session.replace_match(match, "ZZMARKERZZ")
    assert edit.success is True

    assert len(english_session.find_all_by_text("ZZMARKERZZ")) >= 1

    # The replaced span no longer carries the original text at that anchor/offset.
    still_present = any(
        m.enclosing_anchor.id == match.enclosing_anchor.id
        and m.span.start == match.span.start
        and m.text == original
        for m in english_session.grep(r"[A-Za-z]{5,}")
    )
    assert still_present is False
