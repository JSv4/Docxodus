"""Read-only location query and wire-decoding contracts through the real stdio host."""
from __future__ import annotations

import io
import re
import xml.etree.ElementTree as ET
from zipfile import ZipFile

from docx_scalpel import AnchorLocation, AnchorLocationsResult, DocxSessionSettings, open_session


def test_location_wire_keeps_canonical_identity_and_unknown_nulls() -> None:
    result = AnchorLocationsResult._from_wire({
        "version": 7,
        "locations": {"p:body:abcd": {"anchorId": "p:body:abcdef", "partUri": "/word/document.xml",
                                      "elementPath": "/*[1]/*[1]/*[2]"}, "unknown": None},
    })
    assert result.version == 7
    assert result.locations["p:body:abcd"] == AnchorLocation("p:body:abcdef", "/word/document.xml", "/*[1]/*[1]/*[2]")
    assert result.locations["unknown"] is None


def test_locations_resolve_in_original_xml_and_do_not_advance_history(tour_plan_bytes: bytes) -> None:
    with open_session(tour_plan_bytes, DocxSessionSettings(capture_initial_projection=False)) as session:
        result = session.get_anchor_locations()
        assert result.version == session.get_version() == 0
        assert result.locations
        with ZipFile(io.BytesIO(tour_plan_bytes)) as package:
            for location in result.locations.values():
                assert location is not None
                steps = [int(value) for value in re.findall(r"/\*\[(\d+)\]", location.element_path)]
                assert steps[0] == 1
                element = ET.fromstring(package.read(location.part_uri.lstrip("/")))
                for step in steps[1:]:
                    element = list(element)[step - 1]
                assert element is not None
        assert session.get_anchor_locations(None) == result
        assert session.get_version() == 0
        assert session.undo() is False


def test_location_subset_and_empty_are_distinct_from_all(tour_plan_bytes: bytes) -> None:
    with open_session(tour_plan_bytes) as session:
        all_locations = session.get_anchor_locations()
        key = next(iter(all_locations.locations))
        unknown = 'not " an anchor'
        subset = session.get_anchor_locations([key, unknown, key])
        assert list(subset.locations) == [key, unknown]
        assert subset.locations[key] == all_locations.locations[key]
        assert subset.locations[unknown] is None
        assert session.get_anchor_locations([]).locations == {}
