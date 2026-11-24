"""Tab stops: ``get_tab_stops``, ``set_tab_stops`` and ``set_default_tab_stop``."""

from __future__ import annotations

import zipfile
from io import BytesIO
from pathlib import Path

import pytest

from docx_scalpel import TabStop, open_session

FIXTURE = Path(__file__).parents[2] / "TestFiles" / "DA001-TemplateDocument.docx"


def _first_body_anchor(session) -> str:
    for anchor_id in session.project().anchor_index:
        if anchor_id.startswith("p:body:"):
            return anchor_id
    pytest.fail("no body paragraph in fixture")


def _part(saved: bytes, name: str) -> str:
    with zipfile.ZipFile(BytesIO(saved)) as zf:
        return zf.read(name).decode("utf-8")


def test_set_tab_stops_reads_back_saves_and_undoes() -> None:
    with open_session(FIXTURE.read_bytes()) as session:
        anchor = _first_body_anchor(session)
        before = session.get_tab_stops(anchor)
        assert before is not None and before.stops == ()

        result = session.set_tab_stops(anchor, [
            TabStop(1440),
            TabStop(4320, alignment="right", leader="dot"),
        ])

        assert result.success, result.error
        info = session.get_tab_stops(anchor)
        assert info is not None
        assert info.stops == (
            TabStop(1440, "left", "none"),
            TabStop(4320, "right", "dot"),
        )
        xml = _part(session.save(), "word/document.xml")
        assert 'w:val="right"' in xml and 'w:leader="dot"' in xml and 'w:pos="4320"' in xml
        assert session.undo()
        assert session.get_tab_stops(anchor).stops == ()


def test_set_default_tab_stop_writes_settings() -> None:
    with open_session(FIXTURE.read_bytes()) as session:
        anchor = _first_body_anchor(session)
        assert session.set_default_tab_stop(1080).success
        assert session.get_tab_stops(anchor).default_tab_stop_twips == 1080
        assert 'w:defaultTabStop w:val="1080"' in _part(session.save(), "word/settings.xml")


def test_bad_values_are_refused() -> None:
    with open_session(FIXTURE.read_bytes()) as session:
        anchor = _first_body_anchor(session)
        refused = session.set_tab_stops(anchor, [TabStop(40000)])
        assert not refused.success
        assert refused.error.code.value == "invalid_paragraph_format"
        assert not session.set_default_tab_stop(0).success
