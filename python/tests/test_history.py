from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor
from dataclasses import FrozenInstanceError

import pytest

from docx_scalpel import (
    DocxHistoryError, DocxVersionMetadata, HistoryBlobReference, HistoryHead,
    VerificationDigest, convert_docx_to_html, docx_diff_compare_products,
    generate_package_manifest, open_history, open_session, shutdown_host,
)
from docx_scalpel.errors import DocxodusTransportError


def metadata(hour: int = 12) -> DocxVersionMetadata:
    return DocxVersionMetadata("python-actor", f"2026-01-01T{hour:02}:00:00Z", application_metadata={"matter": "123"})


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
