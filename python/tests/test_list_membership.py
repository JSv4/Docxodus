"""``set_list_membership``: a paragraph joins an existing list at a level (#1791), as Word's
Format Painter makes a paragraph painted from a numbered one number on in its list."""

from __future__ import annotations

from pathlib import Path
from typing import Iterator

import pytest

from docx_scalpel import DocxSession, open_session


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
    for anchor in session.project().anchor_index.values():
        if anchor.kind == kind and anchor.scope == "body":
            return anchor
    pytest.skip(f"fixture has no body {kind} anchor")


def test_set_list_membership_joins_the_list_and_undoes(list_session: DocxSession) -> None:
    item = _first_anchor_of_kind(list_session, "li")
    plain = _first_anchor_of_kind(list_session, "p")
    num_id = list_session.get_list_membership(item.id).num_id
    assert list_session.get_list_membership(plain.id) is None

    result = list_session.set_list_membership(plain.id, num_id, 1)

    assert result.success, result.error
    membership = list_session.get_list_membership(result.modified[0].id)
    assert membership is not None
    assert (membership.num_id, membership.level) == (num_id, 1)
    assert list_session.undo()
    assert list_session.get_list_membership(plain.id) is None


def test_set_list_membership_refuses_an_unknown_list(list_session: DocxSession) -> None:
    plain = _first_anchor_of_kind(list_session, "p")

    result = list_session.set_list_membership(plain.id, 99999, 0)

    assert not result.success
    assert result.error.code.value == "invalid_list_level"
    assert list_session.get_list_membership(plain.id) is None
