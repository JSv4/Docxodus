"""Transaction-id retry deduplication through the stdio host (issue #761)."""

from __future__ import annotations

from docx_scalpel import (
    EditErrorCode,
    DocxSession,
    MutationBatchStep,
    open_session,
)


def _first_paragraph(session: DocxSession) -> str:
    projection = session.project()
    return next(
        anchor.id
        for anchor in projection.anchor_index.values()
        if anchor.scope == "body" and anchor.kind in ("p", "h", "li")
    )


def _insert(anchor: str, text: str) -> list[MutationBatchStep]:
    return [MutationBatchStep("insert_paragraph", {
        "anchorId": anchor, "position": "after", "markdown": text,
    })]


def test_identical_retry_replays_the_original_result_without_applying_again(
    tour_plan_bytes: bytes,
) -> None:
    with open_session(tour_plan_bytes) as session:
        anchor = _first_paragraph(session)
        first = session.execute_batch(_insert(anchor, "inserted exactly once"), transaction_id="tx-1")
        version = session.get_version()

        retry = session.execute_batch(_insert(anchor, "inserted exactly once"), transaction_id="tx-1")

        assert first.success and retry == first
        assert first.transaction is not None
        assert first.transaction.transaction_id == "tx-1"
        assert first.transaction.request_fingerprint.startswith("sha256:")
        assert session.get_version() == version
        assert session.project().markdown.count("inserted exactly once") == 1


def test_reusing_an_id_for_a_different_batch_is_a_conflict(tour_plan_bytes: bytes) -> None:
    with open_session(tour_plan_bytes) as session:
        anchor = _first_paragraph(session)
        session.execute_batch(_insert(anchor, "first"), transaction_id="tx-2")
        version = session.get_version()

        conflict = session.execute_batch(_insert(anchor, "second"), transaction_id="tx-2")

        assert not conflict.success
        assert conflict.failure is not None
        assert conflict.failure.error.code is EditErrorCode.TRANSACTION_CONFLICT
        assert conflict.transaction is not None and conflict.transaction.transaction_id == "tx-2"
        assert session.get_version() == version
        assert "second" not in session.project().markdown


def test_a_batch_without_an_id_carries_no_identity_and_is_not_deduplicated(
    tour_plan_bytes: bytes,
) -> None:
    with open_session(tour_plan_bytes) as session:
        anchor = _first_paragraph(session)
        session.execute_batch(_insert(anchor, "twice"))
        again = session.execute_batch(_insert(anchor, "twice"))

        assert again.transaction is None
        assert session.project().markdown.count("twice") == 2


def test_replay_after_undo_returns_the_historical_result_without_redoing(
    tour_plan_bytes: bytes,
) -> None:
    with open_session(tour_plan_bytes) as session:
        anchor = _first_paragraph(session)
        first = session.execute_batch(_insert(anchor, "undone"), transaction_id="tx-3")
        assert session.undo()
        assert "undone" not in session.project().markdown

        replay = session.execute_batch(_insert(anchor, "undone"), transaction_id="tx-3")

        assert replay == first
        assert "undone" not in session.project().markdown
