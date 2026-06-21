"""End-to-end parity test for ``insert_tab`` — Python mirror of the C# DS240-242 tests.

Confirms the ``insert_tab`` op is hooked up through the full stack: docx-scalpel →
NDJSON → ``docxodus-pyhost`` → ``DocxSessionOps.InsertTab`` → ``DocxSession.InsertTab``.
"""

from __future__ import annotations

import re

from docx_scalpel import TabStopAlignment, open_session


def _first_body_paragraph(session) -> str:
    proj = session.project()
    return min(
        (t for t in proj.anchor_index.values() if t.kind in ("p", "h") and t.scope == "body"),
        key=lambda t: (proj.markdown.find("{#" + t.id + "}") + (1 << 30)) % (1 << 31),
    ).id


def test_insert_tab_right_adds_stop_and_tab_run(tour_plan_bytes: bytes) -> None:
    with open_session(tour_plan_bytes) as session:
        anchor = _first_body_paragraph(session)
        result = session.insert_tab(anchor, 0, TabStopAlignment.RIGHT)
        assert result.success, result.error
        xml = session.raw.get_xml(anchor)
        # A right tab STOP on the paragraph (w:tabs/w:tab w:val="right")…
        assert 'w:val="right"' in xml, xml
        # …and a tab RUN in the content: at least two <w:tab> elements (the stop + the run).
        assert len(re.findall(r"<w:tab[ />]", xml)) >= 2, xml


def test_insert_tab_defaults_to_right(tour_plan_bytes: bytes) -> None:
    with open_session(tour_plan_bytes) as session:
        anchor = _first_body_paragraph(session)
        result = session.insert_tab(anchor, 0)  # default alignment = RIGHT
        assert result.success, result.error
        assert 'w:val="right"' in session.raw.get_xml(anchor)
