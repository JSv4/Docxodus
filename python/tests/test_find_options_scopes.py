"""Regression coverage for ``FindOptions`` scope filtering on the Python wrapper.

The .NET ``FindOptions`` record exposes two distinct scope controls:

* ``Scopes`` — a ``ProjectionScopes`` flag set (wire key ``scopes``, an int),
  the coarse "which categories of part" filter (default = All).
* ``ScopeFilter`` — a string naming one specific part such as ``"hdr1"``
  (wire key ``scopeFilter``), a finer post-filter on top of ``Scopes``.

The Python wrapper previously conflated the two: it had a single
``scope_filter: ProjectionScopes`` field and serialized it as an int under the
``scopeFilter`` key. The .NET dispatcher reads ``scopeFilter`` as a *string*
(so the int was dropped) and reads the flag set from ``scopes`` (which the
wrapper never emitted), so scope filtering from Python was a silent no-op.

These tests pin the corrected wire contract and the end-to-end narrowing.
HC031-Complicated-Document.docx carries body text plus footnote/endnote/comment
parts, so restricting ``scopes`` to the body must drop the footnote hit.
"""

from __future__ import annotations

from pathlib import Path
from typing import Iterator

import pytest

from docx_scalpel import DocxSession, open_session
from docx_scalpel.enums import ProjectionScopes
from docx_scalpel.types import FindOptions


@pytest.fixture
def multiscope_session(test_files_dir: Path) -> Iterator[DocxSession]:
    fixture = test_files_dir / "HC031-Complicated-Document.docx"
    if not fixture.exists():
        pytest.skip(f"fixture missing: {fixture}")
    session = open_session(fixture.read_bytes())
    try:
        yield session
    finally:
        session.close()


def test_to_wire_maps_scopes_to_int_and_scope_filter_to_string() -> None:
    combined = ProjectionScopes.HEADERS | ProjectionScopes.FOOTERS
    wire = FindOptions(scopes=combined, scope_filter="hdr1").to_wire()

    # Coarse flag set → numeric "scopes"; named part → string "scopeFilter".
    assert wire["scopes"] == int(combined)
    assert wire["scopeFilter"] == "hdr1"
    # Never the other way around — the int must not leak into "scopeFilter".
    assert not isinstance(wire.get("scopeFilter"), int)


def test_to_wire_omits_unset_scope_fields() -> None:
    wire = FindOptions(ignore_case=True).to_wire()
    assert "scopes" not in wire
    assert "scopeFilter" not in wire


def test_scopes_narrows_find_all_by_text_to_body(multiscope_session: DocxSession) -> None:
    needle = "footnote"  # present in the footnote part of HC031

    all_hits = multiscope_session.find_all_by_text(needle, FindOptions(ignore_case=True))
    if not any(h.scope.startswith("fn") for h in all_hits):
        pytest.skip("fixture footnote text changed; needle no longer in fn scope")

    body_only = multiscope_session.find_all_by_text(
        needle, FindOptions(ignore_case=True, scopes=ProjectionScopes.BODY)
    )

    # Restricting to the body must drop the footnote hit (the regression: this
    # used to be a no-op, so body_only == all_hits).
    assert all(not h.scope.startswith("fn") for h in body_only)
    assert len(body_only) < len(all_hits)


def test_scopes_footnotes_only_returns_footnote_anchor(multiscope_session: DocxSession) -> None:
    fn_only = multiscope_session.find_all_by_text(
        "footnote", FindOptions(ignore_case=True, scopes=ProjectionScopes.FOOTNOTES)
    )
    assert len(fn_only) >= 1
    assert all(h.scope.startswith("fn") for h in fn_only)
