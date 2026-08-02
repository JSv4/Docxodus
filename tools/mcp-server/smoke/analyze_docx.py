#!/usr/bin/env python3
"""Compare package preservation and workflow observables across three DOCX files."""

from __future__ import annotations

import argparse
import hashlib
import json
import zipfile
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
NS = {"w": W, "w15": W15}
WA = f"{{{W}}}"
W15A = f"{{{W15}}}"

MARKERS = [
    "MCP ROUND THREE REVIEW SCHEDULE",
    "DOES HEREBY CERTIFY AND ACKNOWLEDGE:",
    "MCP ROUND THREE CORPORATION",
    "Round-three comment anchor phrase remains visible after review.",
    "Round-three item alpha — text and styles",
    "Round-three item beta — numbering and nesting",
    "Round-three item gamma — comments and revisions",
    "RT3 Feature",
    "Tracked rejection anchor: BASE.",
    "BASE TEMPORARY",
    "PREVIEW-ONLY-MARKER",
]


def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def xml_root(archive: zipfile.ZipFile, name: str) -> ET.Element | None:
    try:
        return ET.fromstring(archive.read(name))
    except KeyError:
        return None


def visible_text(element: ET.Element | None) -> str:
    if element is None:
        return ""
    return "".join(
        node.text or ""
        for node in element.iter()
        if node.tag in {WA + "t", WA + "delText"}
    )


def count(root: ET.Element | None, local_name: str) -> int:
    return 0 if root is None else sum(1 for _ in root.iter(WA + local_name))


def canonical_field_instructions(root: ET.Element | None) -> str:
    if root is None:
        return ""
    values = [node.text or "" for node in root.iter(WA + "instrText")]
    return " ".join(" ".join(values).split())


def inserted_table(document: ET.Element | None) -> dict[str, Any] | None:
    if document is None:
        return None
    table = next(
        (candidate for candidate in document.iter(WA + "tbl") if "RT3 Feature" in visible_text(candidate)),
        None,
    )
    if table is None:
        return None

    rows = table.findall("w:tr", NS)
    header = rows[0] if rows else None
    tr_pr = None if header is None else header.find("w:trPr", NS)
    height = None if tr_pr is None else tr_pr.find("w:trHeight", NS)
    grid = table.find("w:tblGrid", NS)
    borders = table.find("w:tblPr/w:tblBorders", NS)
    header_cells = [] if header is None else header.findall("w:tc", NS)

    return {
        "rows": len(rows),
        "columns": len(header_cells),
        "columnWidthsTwips": []
        if grid is None
        else [int(column.get(WA + "w", "0")) for column in grid.findall("w:gridCol", NS)],
        "repeatHeader": tr_pr is not None and tr_pr.find("w:tblHeader", NS) is not None,
        "allowBreakAcrossPages": not (
            tr_pr is not None and tr_pr.find("w:cantSplit", NS) is not None
        ),
        "heightTwips": None if height is None else int(height.get(WA + "val", "0")),
        "heightRule": None if height is None else height.get(WA + "hRule"),
        "headerFills": [
            cell.find("w:tcPr/w:shd", NS).get(WA + "fill")
            if cell.find("w:tcPr/w:shd", NS) is not None
            else None
            for cell in header_cells
        ],
        "borderStyles": {}
        if borders is None
        else {
            edge: (
                None
                if borders.find(f"w:{edge}", NS) is None
                else borders.find(f"w:{edge}", NS).get(WA + "val")
            )
            for edge in ("top", "left", "bottom", "right", "insideH", "insideV")
        },
    }


def package(path: Path) -> dict[str, Any]:
    raw = path.read_bytes()
    with zipfile.ZipFile(path) as archive:
        file_infos = [info for info in archive.infolist() if not info.is_dir()]
        part_hashes = {info.filename: sha256(archive.read(info.filename)) for info in file_infos}
        part_sizes = {info.filename: info.file_size for info in file_infos}
        document = xml_root(archive, "word/document.xml")
        comments = xml_root(archive, "word/comments.xml")
        comments_ex = xml_root(archive, "word/commentsExtended.xml")
        all_text = visible_text(document)
        field_instructions = canonical_field_instructions(document)

        return {
            "path": str(path),
            "sha256": sha256(raw),
            "compressedBytes": len(raw),
            "fileParts": len(file_infos),
            "uncompressedBytes": sum(info.file_size for info in file_infos),
            "partHashes": part_hashes,
            "partSizes": part_sizes,
            "markers": {marker: all_text.count(marker) for marker in MARKERS},
            "structure": {
                "paragraphs": count(document, "p"),
                "tables": count(document, "tbl"),
                "tableRows": count(document, "tr"),
                "tableCells": count(document, "tc"),
                "numberingProperties": count(document, "numPr"),
                "sections": count(document, "sectPr"),
                "bookmarks": count(document, "bookmarkStart"),
                "fieldCharacters": count(document, "fldChar"),
                "fieldInstructions": count(document, "instrText"),
                "fieldInstructionTokens": len(field_instructions.split()),
                "fieldInstructionCanonicalSha256": sha256(field_instructions.encode()),
                "insertions": count(document, "ins"),
                "deletions": count(document, "del"),
                "footnoteReferences": count(document, "footnoteReference"),
                "headerParts": sum(
                    name.startswith("word/header") and name.endswith(".xml")
                    for name in part_hashes
                ),
                "footerParts": sum(
                    name.startswith("word/footer") and name.endswith(".xml")
                    for name in part_hashes
                ),
            },
            "comments": {
                "count": 0 if comments is None else len(comments.findall("w:comment", NS)),
                "threadEntries": 0
                if comments_ex is None
                else len(comments_ex.findall("w15:commentEx", NS)),
                "resolvedThreadEntries": 0
                if comments_ex is None
                else sum(
                    item.get(W15A + "done") in {"1", "true", "on"}
                    for item in comments_ex.findall("w15:commentEx", NS)
                ),
                "replyEntries": 0
                if comments_ex is None
                else sum(
                    item.get(W15A + "paraIdParent") is not None
                    for item in comments_ex.findall("w15:commentEx", NS)
                ),
            },
            "insertedTable": inserted_table(document),
        }


def preservation(source: dict[str, Any], output: dict[str, Any]) -> dict[str, Any]:
    source_parts = source["partHashes"]
    output_parts = output["partHashes"]
    source_sizes = source["partSizes"]
    output_sizes = output["partSizes"]
    common = sorted(source_parts.keys() & output_parts.keys())
    unchanged = [name for name in common if source_parts[name] == output_parts[name]]
    changed = [name for name in common if source_parts[name] != output_parts[name]]
    return {
        "sourcePartsUnchanged": len(unchanged),
        "sourcePartsChanged": len(changed),
        "changedParts": changed,
        "changedPartSizeDeltas": {
            name: output_sizes[name] - source_sizes[name] for name in changed
        },
        "addedParts": sorted(output_parts.keys() - source_parts.keys()),
        "removedParts": sorted(source_parts.keys() - output_parts.keys()),
    }


def workflow_observables(document: dict[str, Any]) -> dict[str, Any]:
    return {
        "markers": document["markers"],
        "comments": document["comments"],
        "insertedTable": document["insertedTable"],
        "remainingRevisions": document["structure"]["insertions"]
        + document["structure"]["deletions"],
    }


def workflow_equivalent(reference: dict[str, Any], replacement: dict[str, Any]) -> bool:
    left = workflow_observables(reference)
    right = workflow_observables(replacement)
    left_table = dict(left.pop("insertedTable") or {})
    right_table = dict(right.pop("insertedTable") or {})
    left_widths = left_table.pop("columnWidthsTwips", [])
    right_widths = right_table.pop("columnWidthsTwips", [])
    return (
        left == right
        and left_table == right_table
        and len(left_widths) == len(right_widths)
        and all(abs(a - b) <= 5 for a, b in zip(left_widths, right_widths))
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=Path)
    parser.add_argument("reference", type=Path)
    parser.add_argument("replacement", type=Path)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    source = package(args.source)
    reference = package(args.reference)
    replacement = package(args.replacement)
    source.pop("partHashes")
    source.pop("partSizes")
    reference_hashes = reference["partHashes"]
    reference_sizes = reference["partSizes"]
    replacement_hashes = replacement["partHashes"]
    replacement_sizes = replacement["partSizes"]

    # Re-open the source hashes for preservation calculations after keeping the public
    # report compact and free of one hash per package part.
    source_with_hashes = package(args.source)
    report = {
        "source": source,
        "reference": {
            key: value for key, value in reference.items() if key not in {"partHashes", "partSizes"}
        },
        "replacement": {
            key: value for key, value in replacement.items() if key not in {"partHashes", "partSizes"}
        },
        "preservation": {
            "reference": preservation(
                source_with_hashes,
                {"partHashes": reference_hashes, "partSizes": reference_sizes},
            ),
            "replacement": preservation(
                source_with_hashes,
                {"partHashes": replacement_hashes, "partSizes": replacement_sizes},
            ),
        },
        "workflowObservablesExact": workflow_observables(reference)
        == workflow_observables(replacement),
        "workflowObservablesEquivalent": workflow_equivalent(reference, replacement),
    }
    print(json.dumps(report, indent=2, sort_keys=True))


if __name__ == "__main__":
    main()
