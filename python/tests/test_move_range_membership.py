"""Native host resolution agrees with captured Word move endpoints (#1340)."""
import hashlib
import io
import json
from pathlib import Path
import xml.etree.ElementTree as ET
import zipfile

import pytest
from docx_scalpel import open_session, docx_diff_accept_revisions, docx_diff_reject_revisions

ROOT = Path(__file__).resolve().parents[2] / "TestFiles/RP/MoveRangeMembership"
CASES = json.loads((ROOT / "word-endpoints.json").read_text())["cases"]
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


@pytest.mark.parametrize("sample", CASES, ids=lambda sample: sample["id"])
@pytest.mark.parametrize("accept", [False, True])
@pytest.mark.parametrize("route", ["stateless", "bulk", "selective"])
def test_native_host_matches_word(sample, accept, route):
    source = (ROOT / (sample["id"] + ".docx")).read_bytes()
    assert hashlib.sha256(source).hexdigest() == sample["inputSha256"]
    if route == "stateless":
        output = (docx_diff_accept_revisions if accept else docx_diff_reject_revisions)(source)
    else:
        with open_session(source) as session:
            moves = session.list_revisions()
            assert len(moves) == 1 and moves[0].type == "move"
            assert moves[0].resolution_status == "supported"
            if route == "selective":
                result = (session.accept_revision if accept else session.reject_revision)(moves[0].id)
            else:
                result = (session.accept_all_revisions if accept else session.reject_all_revisions)()
            assert result.success, result.error
            assert not session.list_revisions()
            output = session.save()
    with zipfile.ZipFile(io.BytesIO(output)) as package:
        xml = ET.fromstring(package.read("word/document.xml"))
    paragraphs = ["".join(node.text or "" for node in p.iter(W + "t")) for p in xml.find(W + "body").findall(W + "p")]
    assert paragraphs == sample["word"]["accept-all" if accept else "reject-all"]["paragraphs"]
    markers = {W + kind for kind in ("moveFrom", "moveTo", "moveFromRangeStart", "moveFromRangeEnd", "moveToRangeStart", "moveToRangeEnd")}
    assert not any(node.tag in markers for node in xml.iter())
