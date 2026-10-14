"""Consume actual browser outputs in fresh native hosts; never regenerate the source corpus."""

from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path
import xml.etree.ElementTree as ET
from zipfile import ZipFile

import pytest

from docx_scalpel import (
    DocxHistoryError, DocxVersionMetadata, generate_package_manifest,
    open_history, open_history_archive, open_session, shutdown_host,
)


def test_browser_files_reopen_import_continue_and_recover_in_native_hosts(tmp_path, test_files_dir):
    source_directory = os.environ.get("DOCXODUS_HISTORY_BROWSER_ARTIFACT_DIR")
    if source_directory is None:
        pytest.skip("Run after Playwright with DOCXODUS_HISTORY_BROWSER_ARTIFACT_DIR set")
    archives = sorted(Path(source_directory).rglob("browser-continued.docxhistory"))
    assert archives, "Configured browser artifact directory contains no history-file outputs"
    fixture = test_files_dir / "HistoryArchive"
    index = json.loads((fixture / "agreement.json").read_text())
    with open_history_archive((fixture / "agreement.docxhistory").read_bytes()) as original_reader:
        original_view = original_reader.read()

    for number, archive_path in enumerate(archives):
        output = tmp_path / f"handoff-{number}"
        output.mkdir()
        paths = [archive_path, archive_path.with_name("browser-latest.docx"),
                 archive_path.with_name("browser-arbitrary-comparison.docx")]
        source, latest, browser_redline = (path.read_bytes() for path in paths)
        original_hashes = {str(path): hashlib.sha256(path.read_bytes()).hexdigest() for path in paths}
        shutdown_host()  # Do not inherit any client/handle/store from another test.
        with open_history_archive(source) as reader:
            view = reader.read()
            assert view.head.revision == index["head"]["revision"] + 2
            assert reader.document_id == index["documentId"]
            page = reader.list_versions(limit=100)
            assert page.next is None and len(page.versions) == len(index["versions"]) + 2
            for expected in index["versions"]:
                version = next(v for v in page.versions if v.id.to_wire() == expected["id"])
                exact = reader.export_docx(version.id)
                assert exact == (fixture / expected["file"]).read_bytes()
                assert hashlib.sha256(exact).hexdigest() == expected["sha256"]
            first = page.versions[-1]
            revision = next(v for v in page.versions if v.record.metadata.label == "Counsel revision")
            assert reader.export_docx() == latest == reader.export_docx(revision.id)
            assert reader.materialize(view.state.sequence) == latest
            assert generate_package_manifest(reader.replay(view.state.sequence)).ordered_opc_content_digest == \
                generate_package_manifest(latest).ordered_opc_content_digest
            native_redline = reader.compare_versions(first.id, revision.id)
            # ZIP serialization is not the interoperability contract; compare all OPC content.
            assert generate_package_manifest(native_redline).ordered_opc_content_digest == \
                generate_package_manifest(browser_redline).ordered_opc_content_digest
            (output / "native-arbitrary-comparison.docx").write_bytes(native_redline)

        with ZipFile(paths[2]) as redline_zip:
            document = ET.fromstring(redline_zip.read("word/document.xml"))
            w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
            assert list(document.iter(w + "ins")) or list(document.iter(w + "del"))
        with ZipFile(paths[1]) as latest_zip:
            assert "history.json" not in latest_zip.namelist()
            assert not any(name.startswith("blobs/") for name in latest_zip.namelist())

        with open_history(output / "store") as history:
            imported = history.import_history_archive(source)
            assert imported.view == view and not imported.already_present
            assert history.import_history_archive(source).already_present
        shutdown_host()
        initial_bytes = (fixture / "agreement-v1.docx").read_bytes()
        save_metadata = DocxVersionMetadata("native-handoff", "2026-09-02T12:00:00Z", label="Native checkpoint")
        with open_history(output / "store") as history:
            doc = history.document(index["documentId"])
            assert doc.read() == view
            # The receipt originated before the browser import; handoff must retain it exactly.
            original = doc.create_version(None, initial_bytes, first.record.metadata, request_id="initial")
            assert original.head.revision == 1 and original.version.id == first.id
            # These two receipts were written by WASM, not Python. Reuse the browser's exact
            # original head/bytes/metadata/target and require its original results after import.
            browser_saved = doc.create_version(original_view.head, initial_bytes,
                DocxVersionMetadata("frontend", "2026-09-01T12:00:00Z", label="Browser checkpoint"),
                request_id="browser-checkpoint")
            assert browser_saved.head.revision == original_view.head.revision + 1
            assert browser_saved.version == next(v for v in page.versions if v.record.metadata.label == "Browser checkpoint")
            browser_restored = doc.restore_version(browser_saved.head, revision.id,
                DocxVersionMetadata("frontend", "2026-09-01T13:00:00Z"), request_id="browser-restore")
            assert browser_restored == view and doc.read() == view
            saved = doc.create_version(view.head, initial_bytes, save_metadata, request_id="native-checkpoint")
            assert saved.head.revision == view.head.revision + 1
        shutdown_host()  # Retry after a real process restart, then append a restore.
        with open_history(output / "store") as history:
            doc = history.document(index["documentId"])
            assert doc.create_version(view.head, initial_bytes, save_metadata, request_id="native-checkpoint") == saved
            restored = doc.restore_version(saved.head, revision.id,
                DocxVersionMetadata("native-handoff", "2026-09-02T13:00:00Z"), request_id="native-restore")
            assert restored.head.revision == saved.head.revision + 1
            with pytest.raises(DocxHistoryError) as conflict:
                history.import_history_archive(source)
            assert conflict.value.code == "ImportConflict"
            (output / "native-continued.docxhistory").write_bytes(doc.export_history_archive())
            (output / "native-latest.docx").write_bytes(doc.export_docx())
        shutdown_host()
        with open_history_archive((output / "native-continued.docxhistory").read_bytes()) as reader:
            assert reader.read() == restored
            assert reader.export_docx() == latest == (output / "native-latest.docx").read_bytes()
            final_page = reader.list_versions(limit=100)
            assert len(final_page.versions) == len(page.versions) + 2
            with open_session(reader.export_docx()) as session:
                assert session.project().markdown
            (output / "native-standalone.docxhistory").write_bytes(reader.export_history_archive())
        shutdown_host()
        with open_history_archive((output / "native-standalone.docxhistory").read_bytes()) as reader:
            assert reader.read() == restored and reader.list_versions(limit=100) == final_page
            assert reader.export_docx() == latest
        shutdown_host()
        assert original_hashes == {str(path): hashlib.sha256(path.read_bytes()).hexdigest() for path in paths}
        report = {"sourceFiles": original_hashes, "browserHead": view.head.to_wire(),
                  "nativeHead": restored.head.to_wire(), "retainedVersionIds": [v.id.to_wire() for v in page.versions],
                  "outputFiles": {path.name: hashlib.sha256(path.read_bytes()).hexdigest()
                                  for path in output.iterdir() if path.is_file()}}
        (output / "handoff.json").write_text(json.dumps(report, indent=2) + "\n")
        print(f"Verified browser/native history artifacts: {output}")
