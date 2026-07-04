"""Accept/reject round-trip over the ``docx-scalpel`` DocxDiff client surface.

These are deliberately NOT shape/length checks. They drive the full client wire —
``docx_diff_compare`` (bytes out) → ``docx_diff_accept_revisions`` /
``docx_diff_reject_revisions`` (bytes in→out, the new surface) → ``convert_docx_to_html``
(for a text-level oracle) — and assert the engine's round-trip contract: accepting a
redline's revisions reproduces the *right* document and rejecting them reproduces the
*left*, at the per-block text level. A wire/type-mapping break anywhere in that diff
path (a mis-encoded blob, a dropped ``docxB64`` field, a base64 corruption) changes the
bytes that come back and breaks the text equality below.
"""

from __future__ import annotations

import re
from pathlib import Path

import pytest

from docx_scalpel import (
    convert_docx_to_html,
    docx_diff_accept_revisions,
    docx_diff_compare,
    docx_diff_reject_revisions,
)

# (left, right) WC pairs whose edits land in body text, so the HTML-projection oracle
# sees a genuine difference between the two sides.
PAIRS = [
    ("WC/WC001-Digits.docx", "WC/WC001-Digits-Mod.docx"),
    ("WC/WC004-Large.docx", "WC/WC004-Large-Mod.docx"),
]


def _text(docx: bytes) -> str:
    """Visible text of a DOCX via the HTML projection: tags stripped, whitespace folded."""
    html = convert_docx_to_html(docx)
    txt = re.sub(r"<[^>]+>", " ", html)
    txt = txt.replace("&nbsp;", " ").replace("&amp;", "&")
    return re.sub(r"\s+", " ", txt).strip()


@pytest.mark.parametrize("left_rel,right_rel", PAIRS)
def test_accept_reject_round_trip(test_files_dir: Path, left_rel: str, right_rel: str) -> None:
    left_path, right_path = test_files_dir / left_rel, test_files_dir / right_rel
    if not left_path.exists() or not right_path.exists():
        pytest.skip(f"fixture absent: {left_rel} / {right_rel}")
    left, right = left_path.read_bytes(), right_path.read_bytes()

    redline = docx_diff_compare(left, right)
    assert isinstance(redline, bytes) and len(redline) > 1000

    accepted = docx_diff_accept_revisions(redline)
    rejected = docx_diff_reject_revisions(redline)
    assert isinstance(accepted, bytes) and len(accepted) > 1000
    assert isinstance(rejected, bytes) and len(rejected) > 1000

    left_text, right_text = _text(left), _text(right)
    assert left_text != right_text, "fixtures must genuinely differ for the round-trip to mean anything"

    # The contract: accept materializes the right side, reject the left.
    assert _text(accepted) == right_text
    assert _text(rejected) == left_text


def test_accept_reject_are_distinct(test_files_dir: Path) -> None:
    """Accept and reject of the same redline produce DIFFERENT documents — guards against a
    no-op/passthrough wire bug where both paths echo their input unchanged."""
    left = (test_files_dir / "WC" / "WC001-Digits.docx").read_bytes()
    right = (test_files_dir / "WC" / "WC001-Digits-Mod.docx").read_bytes()
    redline = docx_diff_compare(left, right)
    assert _text(docx_diff_accept_revisions(redline)) != _text(docx_diff_reject_revisions(redline))


def test_docx_diff_settings_track_block_format_changes_to_wire() -> None:
    """The track_block_format_changes opt-out only emits its wire key when disabled
    (default True → omitted, matching the host's default-on behavior)."""
    from docx_scalpel.types import DocxDiffSettings

    assert DocxDiffSettings().track_block_format_changes is True
    assert "trackBlockFormatChanges" not in DocxDiffSettings().to_wire()
    assert DocxDiffSettings(track_block_format_changes=False).to_wire()[
        "trackBlockFormatChanges"
    ] is False
