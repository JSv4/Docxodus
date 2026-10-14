from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor
from dataclasses import FrozenInstanceError

import pytest

from docx_scalpel import (
    DocxHistoryError, DocxVersionMetadata, HistoryBlobReference, HistoryHead,
    VerificationDigest, convert_docx_to_html, docx_diff_compare_products,
    generate_package_manifest, open_history, open_history_archive, open_session, shutdown_host,
)
from docx_scalpel.errors import DocxodusTransportError


def metadata(hour: int = 12) -> DocxVersionMetadata:
    return DocxVersionMetadata("python-actor", f"2026-01-01T{hour:02}:00:00Z", application_metadata={"matter": "123"})


@pytest.mark.parametrize("name", ["agreement", "charter-collaboration"])
def test_real_portable_files_drive_readonly_and_writable_controls_after_process_restart(tmp_path, test_files_dir, name):
    root = test_files_dir / "HistoryArchive"
    source = (root / f"{name}.docxhistory").read_bytes()
    with open_history_archive(source) as reader:
        info = reader.info
        assert not hasattr(reader, "create_version") and not hasattr(reader, "restore_version")
        page = reader.list_versions(limit=100)
        first = page.versions[-1]
        expected = (root / f"{name}-v1.docx").read_bytes()
        assert reader.export_docx(first.id) == expected
        if name == "charter-collaboration":
            operations = reader.read_operations_since(None)
            conflict = next(op for op in operations.operations if op.record.status == "conflict")
            assert reader.get_operation(conflict.id) == conflict
            assert reader.export_operation_proposal(conflict.id) == (root / f"{name}-conflicting-proposal.docx").read_bytes()
            assert any(op.input.request.resolves == conflict.id and op.record.status == "accepted" for op in operations.operations)
            assert reader.read_operations_since(operations.view.head).operations == ()
        else:
            redline = reader.compare_versions(first.id, page.versions[-2].id)
            (tmp_path / "arbitrary-comparison.docx").write_bytes(redline)
            with open_session(redline) as compared:
                assert compared.project().markdown
        with pytest.raises(DocxHistoryError) as readonly:
            reader._client.create_version(info.document_id, None, expected, metadata())
        assert readonly.value.code == "ReadOnly"  # Enforced beneath the wrapper too.
    with pytest.raises(DocxHistoryError) as closed:
        reader.read()
    assert closed.value.code == "Closed"
    with open_history(tmp_path / "store") as history:
        imported = history.import_history_archive(source)
        assert imported.archive == info and not imported.already_present
        assert history.import_history_archive(source).already_present
    shutdown_host()
    with open_history(tmp_path / "store") as history:
        doc = history.document(info.document_id)
        with pytest.raises(AttributeError):
            doc.document_id = "other"
        assert doc.document_id == info.document_id
        assert doc.read().head == info.head
        if name == "agreement":
            retry = doc.create_version(None, expected, first.record.metadata, request_id="initial")
            assert retry.head.revision == 1
        saved = doc.create_version(info.head, expected, metadata(), request_id="python-checkpoint")
        assert saved.head.revision == info.head.revision + 1
        restored = doc.restore_version(saved.head, page.versions[0].id, metadata(13), request_id="python-restore")
        assert restored.head.revision == saved.head.revision + 1
        with pytest.raises(DocxHistoryError) as conflict:
            history.import_history_archive(source)
        assert conflict.value.code == "ImportConflict"
        portable = doc.export_history_archive()
        (tmp_path / "continued.docxhistory").write_bytes(portable)
        (tmp_path / "latest.docx").write_bytes(doc.export_docx())
    shutdown_host()
    with open_history_archive((tmp_path / "continued.docxhistory").read_bytes()) as reader:
        assert reader.read().head == restored.head
        assert reader.export_docx() == (tmp_path / "latest.docx").read_bytes()
        with open_session(reader.export_docx()) as session:
            assert session.project().markdown
    assert (root / f"{name}.docxhistory").read_bytes() == source


def test_malformed_portable_file_has_typed_error():
    with pytest.raises(DocxHistoryError) as error:
        open_history_archive(b"not a ZIP")
    assert error.value.code == "InvalidManifest"


def test_durable_request_ids_survive_process_restart_and_return_original_results(tmp_path, tour_plan_bytes):
    with open_history(tmp_path) as history:
        first = history.create_version("doc", None, tour_plan_bytes, metadata(), request_id="first")
        restored = history.restore_version("doc", first.head, first.version.id, metadata(13), request_id="restore")
        later = history.create_version("doc", restored.head, tour_plan_bytes, metadata(14))
        assert first.state.requests.current.id == "first"
        assert restored.state.requests.current.id == "restore"
        assert later.state.requests.current is None and later.state.requests.index is not None
        assert isinstance(restored.state.requests.revision, int)
        with pytest.raises(DocxHistoryError) as conflict:
            history.create_version("doc", later.head, tour_plan_bytes, metadata(), request_id="first")
        assert conflict.value.code == "RequestConflict"
        with pytest.raises(DocxHistoryError) as invalid:
            history.create_version("new", None, tour_plan_bytes, metadata(), request_id=" ")
        assert invalid.value.code == "InvalidRequest"
        assert history.read("new") is None
    shutdown_host()
    with open_history(tmp_path) as reopened:
        assert reopened.create_version("doc", None, tour_plan_bytes, metadata(), request_id="first").head == first.head
        assert reopened.restore_version("doc", first.head, first.version.id, metadata(13), request_id="restore").head == restored.head
        assert reopened.read("doc").head == later.head
        assert len(reopened.list_versions("doc").versions) == 3
        assert reopened.export_version("doc", first.version.id) == tour_plan_bytes


def test_file_history_exact_reopen_time_render_restore_and_compare(tmp_path, tour_plan_bytes):
    root = tmp_path / "history"
    with open_history(root) as a, open_history(root) as b:
        first = a.create_version("doc", None, tour_plan_bytes, metadata())
        assert first.state.sequence == 0
        assert isinstance(first.head.revision, int)
        with open_session(tour_plan_bytes) as session:
            anchor = next(key for key in session.project().anchor_index if key.startswith("p:body:"))
            assert session.replace_text(anchor, "Shared Python historical content.").success
            edited = session.save()
        second = b.create_version("doc", first.head, edited, metadata(13))
        sequence = a.resolve_sequence_at_time("doc", "2026-01-01T13:00:00Z")
        assert sequence == 1
        replay = a.replay("doc", sequence)
        assert generate_package_manifest(replay).ordered_opc_content_digest == generate_package_manifest(edited).ordered_opc_content_digest
        assert "Shared Python historical content." in convert_docx_to_html(a.materialize("doc", sequence))
        comparison = docx_diff_compare_products(a.export_version("doc", first.version.id), a.export_version("doc", second.version.id))
        assert comparison.revisions
        first_page = a.list_versions("doc", limit=1)
        assert first_page.versions[0].id == second.version.id
        assert a.list_versions("doc", first_page.next).versions[0].id == first.version.id
        assert a.get_version("doc", first.version.id).record.metadata.application_metadata["matter"] == "123"
        with pytest.raises(DocxHistoryError) as stale:
            a.create_version("doc", first.head, tour_plan_bytes, metadata())
        assert stale.value.code == "StaleHead"
        restored = b.restore_version("doc", second.head, first.version.id, metadata(14))
        assert restored.state.sequence == 2 and restored.state.epoch == 1
        assert restored.version.record.restored_from == first.version.id
        assert a.export_version("doc", second.version.id) == edited
        updates = a.read_changes_since("doc", first.head)
        assert [entry.commit.sequence for entry in updates.entries] == [1, 2]
        assert updates.entries[0].id == second.state.commit
        assert updates.entries[1].metadata.author == "python-actor"
        assert updates.reset and updates.view.head == restored.head
        assert a.read_changes_since("doc", restored.head).entries == ()

    shutdown_host()
    with open_history(root) as reopened:
        assert reopened.read("doc").head == restored.head
        assert reopened.export_version("doc", restored.version.id) == tour_plan_bytes
        assert reopened.export_version("doc", first.version.id) == tour_plan_bytes


def test_memory_history_lifecycle_and_handles_do_not_attach_to_restarted_host(tour_plan_bytes):
    old = open_history()
    old.create_version("doc", None, tour_plan_bytes, metadata())
    shutdown_host()
    with open_history() as new:
        assert new.read("doc") is None
        with pytest.raises(DocxodusTransportError):
            old.read("doc")
    new.close()
    with pytest.raises(DocxHistoryError) as closed:
        new.read("doc")
    assert closed.value.code == "Closed"


def test_two_file_clients_cannot_overwrite_same_expected_head(tmp_path, tour_plan_bytes):
    with open_history(tmp_path) as a, open_history(tmp_path) as b:
        first = a.create_version("doc", None, tour_plan_bytes, metadata())

        def publish(client):
            try:
                return client.create_version("doc", first.head, tour_plan_bytes, metadata(13))
            except DocxHistoryError as error:
                assert error.code == "StaleHead"
                return None

        with ThreadPoolExecutor(max_workers=2) as pool:
            results = list(pool.map(publish, [a, b]))
        assert sum(result is not None for result in results) == 1
        assert a.read("doc").head.revision == 2


def test_value_types_copy_metadata_and_keep_exact_int64_positions():
    values = {"key": "before"}
    record = DocxVersionMetadata("author", "2026-01-01T00:00:00Z", application_metadata=values)
    values["key"] = "after"
    assert record.application_metadata["key"] == "before"
    with pytest.raises(TypeError):
        record.application_metadata["key"] = "mutation"
    with pytest.raises(FrozenInstanceError):
        record.author = "mutation"
    reference = HistoryBlobReference(VerificationDigest("SHA-256", "a" * 64), 3)
    head = HistoryHead(9223372036854775807, reference)
    assert head.to_wire()["revision"] == "9223372036854775807"
    assert HistoryHead._from_wire(head.to_wire()) == head
    for invalid in [-1, 9223372036854775808, True, 1.5]:
        with pytest.raises(ValueError):
            HistoryHead(invalid, reference).to_wire()
