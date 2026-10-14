#!/usr/bin/env python3
"""Check generated tracked-drawing fixtures against the measured LibreOffice baseline.

Generate fixtures with tracked_drawings_fixture.py first. This runner verifies schema reports,
source hashes, fresh conversion results and the known failure pattern. It writes every observation,
including failures, to results.json. Unknown renderer versions are reported with exit code 2.
"""
import argparse
import hashlib
import json
from pathlib import Path
import re

from lo_render import render

MEASURED_VERSIONS = {"24.2.7.2", "25.8.7.3"}


def fixtures(folder):
    matrix = json.loads((folder / "matrix.json").read_text())
    pairs = json.loads((folder / "pairs.json").read_text())
    if len(matrix) != 48 or len({row["name"] for row in matrix}) != 48:
        raise ValueError("The isolation matrix must contain 48 distinct fixtures")
    required_pairs = {"header-" + revision + "-" + kind
                      for kind in ["group", "vml"] for revision in ["insert", "delete"]}
    if len(pairs) != 4 or {pair["name"] for pair in pairs} != required_pairs:
        raise ValueError("The four insertion/deletion comparison controls are required")
    inputs = {row["name"] for row in matrix}
    inputs.update(pair[side] for pair in pairs for side in ["left", "right"])
    inputs.update("header-" + kind + "-" + revision + ".docx"
                  for kind in ["group", "vml"] for revision in ["ins", "del"])
    diagnostics = {"header-group-paragraph-mark-" + revision for revision in ["ins", "del"]}
    inputs.update(name + ".docx" for name in diagnostics)
    outputs = {pair["name"] + suffix + ".docx" for pair in pairs
               for suffix in ["", "-accepted", "-rejected"]}
    outputs.update(name + suffix + ".docx" for name in diagnostics for suffix in ["-accepted", "-rejected"])
    reports = json.loads((folder / "input-validation.json").read_text())
    for result in json.loads((folder / "results.json").read_text()):
        if "Error" in result:
            raise ValueError("Fixture comparison failed: " + result["Name"])
        if result["Name"] in required_pairs:
            for view, source in [("Accepted", "Right"), ("Rejected", "Left")]:
                if any(result[view][kind] != result[source][kind] for kind in ["Drawings", "Pictures"]):
                    raise ValueError(result["Name"] + ": accept/reject content differs from input")
        reports.extend(value for value in result.values() if isinstance(value, dict) and "File" in value)
    inspected = {report["File"]: report for report in reports}
    files = [folder / name for name in sorted(inputs)] + [folder / "generated" / name for name in sorted(outputs)]
    for file in files:
        report = inspected.get(file.name)
        if report is None or report["Errors"]:
            raise ValueError("Missing or failing schema validation for " + file.name)
        if report["Sha256"] != hashlib.sha256(file.read_bytes()).hexdigest():
            raise ValueError("Stale validation report for " + file.name)
    known_failures = {row["name"] for row in matrix if row["grouped"] and row["textbox"] and
                      row["anchored"] and row["header"] and row["revision"] != "plain"}
    known_failures.update(["header-group-ins.docx", "header-group-del.docx",
                           "header-insert-group.docx", "header-delete-group.docx"])
    return files, known_failures


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("fixtures", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--soffice", default="soffice")
    args = parser.parse_args()
    files, known_failures = fixtures(args.fixtures.resolve())
    args.output.mkdir(parents=True, exist_ok=True)
    observations = []
    for file in files:
        result = render(file, args.output / "pdf", args.soffice)
        result["file"] = file.name
        result["expected_failure"] = file.name in known_failures
        observations.append(result)
        print(file.name + ": " + result["status"], flush=True)
    versions = {match.group(1) for result in observations
                if (match := re.search(r"LibreOffice\s+(\d+(?:\.\d+){3})", result["renderer_version"] or ""))}
    baseline_checked = len(versions) == 1 and versions <= MEASURED_VERSIONS
    failures = []
    for result in observations:
        if not result["source_preserved"]:
            failures.append(result["file"] + ": source changed")
        elif result["expected_failure"]:
            if baseline_checked and (result["failure_code"] != "renderer_failed" or
                                     "Unspecified Application Error" not in result["stderr"]):
                failures.append(result["file"] + ": measured failure pattern changed")
        elif result["status"] != "rendered":
            failures.append(result["file"] + ": control failed to render")
    report = dict(baseline_checked=baseline_checked, versions=sorted(versions), failures=failures,
                  rendered=sum(result["status"] == "rendered" for result in observations),
                  total=len(observations), observations=observations)
    (args.output / "results.json").write_text(json.dumps(report, indent=2))
    print(json.dumps({key: value for key, value in report.items() if key != "observations"}, indent=2))
    return 1 if failures else (0 if baseline_checked else 2)


if __name__ == "__main__":
    raise SystemExit(main())
