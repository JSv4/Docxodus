#!/usr/bin/env python3
"""Render a DOCX with isolated LibreOffice state and report failures without changing it.

Usage: lo_render.py INPUT.docx OUTPUT_DIR [--soffice PATH] [--timeout SECONDS]
The JSON report describes this conversion, not general renderer support or visual fidelity.
"""
import argparse
import hashlib
from io import BytesIO
import json
import os
from pathlib import Path
import signal
import subprocess
import tempfile
import xml.etree.ElementTree as ET
from zipfile import BadZipFile, ZipFile

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
WPG = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"


def _execute(arguments, timeout):
    process = subprocess.Popen(arguments, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                               text=True, start_new_session=True,
                               env=dict(os.environ, SAL_USE_VCLPLUGIN="svp"))
    timed_out = False
    try:
        stdout, stderr = process.communicate(timeout=timeout)
    except subprocess.TimeoutExpired:
        timed_out = True
        # soffice may launch soffice.bin. Terminate this isolated process group so an import
        # that hangs cannot leave a renderer running after its profile has been removed.
        try:
            os.killpg(process.pid, signal.SIGTERM)
        except ProcessLookupError:
            pass
        try:
            stdout, stderr = process.communicate(timeout=2)
        except subprocess.TimeoutExpired:
            try:
                os.killpg(process.pid, signal.SIGKILL)
            except ProcessLookupError:
                pass
            stdout, stderr = process.communicate(timeout=2)
    return dict(returncode=process.returncode, stdout=stdout, stderr=stderr, timed_out=timed_out)


def tracked_drawings(package_bytes):
    """Describe observed revision/drawing combinations; no predictive compatibility rejection."""
    drawings = []
    warnings = []
    revision_names = {"{" + W + "}" + name for name in ["ins", "del", "moveFrom", "moveTo"]}
    try:
        with ZipFile(BytesIO(package_bytes)) as package:
            for part in package.namelist():
                if not part.endswith(".xml"):
                    continue
                try:
                    root = ET.fromstring(package.read(part))
                except ET.ParseError:
                    warnings.append("Could not inspect XML part " + part)
                    continue
                for revision in root.iter():
                    if revision.tag not in revision_names:
                        continue
                    for drawing in revision.iter("{" + W + "}drawing"):
                        drawings.append(dict(
                            part=part, revision=revision.tag.rsplit("}", 1)[-1],
                            anchored=drawing.find(".//{" + WP + "}anchor") is not None,
                            grouped=drawing.find(".//{" + WPG + "}wgp") is not None,
                            textbox=drawing.find(".//{" + WPS + "}txbx") is not None))
    except BadZipFile:
        warnings.append("Input is not a readable ZIP package")
    return drawings, warnings


def render(source, output_dir, executable="soffice", timeout=30):
    source = Path(source).resolve()
    output_dir = Path(output_dir).resolve()
    report = dict(source=str(source), source_sha256=None, source_preserved=True,
                  renderer=str(executable), renderer_version=None, status="failed",
                  failure_code=None, returncode=None, stdout="", stderr="", pdf=None,
                  tracked_drawings=[], inspection_warnings=[])
    try:
        original = source.read_bytes()
    except OSError as error:
        report.update(failure_code="input_unreadable", stderr=str(error))
        return report
    report["source_sha256"] = hashlib.sha256(original).hexdigest()
    report["tracked_drawings"], report["inspection_warnings"] = tracked_drawings(original)
    stage = "renderer"
    try:
        version = _execute([str(executable), "--version"], min(timeout, 10))
        report["renderer_version"] = version["stdout"].strip() or "unknown"
        if version["timed_out"]:
            report.update(failure_code="renderer_timeout", stderr=version["stderr"])
            return report
        stage = "workspace"
        with tempfile.TemporaryDirectory(prefix="docxodus-lo-render-") as folder:
            folder = Path(folder)
            input_dir = folder / "input"
            input_dir.mkdir()
            copied_source = input_dir / source.name
            copied_source.write_bytes(original)
            fresh_output = folder / "output"
            fresh_output.mkdir()
            stage = "renderer"
            result = _execute([
                str(executable), "-env:UserInstallation=" + (folder / "profile").as_uri(),
                "--headless", "--norestore", "--convert-to", "pdf", "--outdir", str(fresh_output),
                str(copied_source)], timeout)
            stage = "workspace"
            report.update(returncode=result["returncode"], stdout=result["stdout"], stderr=result["stderr"])
            pdf = fresh_output / (source.stem + ".pdf")
            if result["timed_out"]:
                report["failure_code"] = "renderer_timeout"
            elif result["returncode"] != 0:
                report["failure_code"] = "renderer_failed"
            elif not pdf.is_file():
                report["failure_code"] = "missing_pdf"
            elif not pdf.read_bytes().startswith(b"%PDF-"):
                report["failure_code"] = "invalid_pdf"
            else:
                try:
                    output_dir.mkdir(parents=True, exist_ok=True)
                    destination = output_dir / pdf.name
                    if destination == source:
                        raise OSError("PDF destination would replace the source")
                    # Replace the directory entry instead of following an existing symlink or
                    # writing through a hard link to the source DOCX.
                    with tempfile.TemporaryDirectory(prefix=".docxodus-pdf-", dir=output_dir) as staging:
                        staged_pdf = Path(staging) / pdf.name
                        staged_pdf.write_bytes(pdf.read_bytes())
                        staged_pdf.replace(destination)
                    report.update(status="rendered", pdf=str(destination))
                except OSError as error:
                    report.update(failure_code="output_unwritable", stderr=str(error))
    except OSError as error:
        report.update(status="unavailable" if stage == "renderer" else "failed",
                      failure_code="renderer_unavailable" if stage == "renderer" else "render_workspace_failed",
                      stderr=str(error))
    finally:
        try:
            report["source_preserved"] = source.read_bytes() == original
        except OSError:
            report["source_preserved"] = False
        if not report["source_preserved"]:
            report.update(status="failed", failure_code="source_changed")
    return report


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=Path)
    parser.add_argument("output_dir", type=Path)
    parser.add_argument("--soffice", default="soffice")
    parser.add_argument("--timeout", type=float, default=30)
    args = parser.parse_args()
    if args.timeout <= 0:
        parser.error("--timeout must be positive")
    report = render(args.source, args.output_dir, args.soffice, args.timeout)
    print(json.dumps(report, indent=2))
    return 0 if report["status"] == "rendered" else 1


if __name__ == "__main__":
    raise SystemExit(main())
