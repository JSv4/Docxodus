"""Guarded commit of a retained preview through the stdio host (issue #760)."""

from __future__ import annotations

from docx_scalpel import (
    DocxSession,
    EditErrorCode,
    MutationBatchStep,
    MutationPreviewRetention,
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


def test_retained_preview_commits_exactly_as_previewed(tour_plan_bytes: bytes) -> None:
    with open_session(tour_plan_bytes) as session:
        anchor = _first_paragraph(session)
        base_version = session.get_version()

        preview = session.preview_batch(_insert(anchor, "committed as previewed"), retain=True)

        assert preview.success and preview.preview
        assert isinstance(preview.retention, MutationPreviewRetention)
        assert preview.retention.base_version == base_version
        assert preview.retention.expires_at.endswith("Z")
        assert session.get_version() == base_version
        assert "committed as previewed" not in session.project().markdown
        created = preview.steps[0].results[0].created[0].id

        commit = session.commit_preview(preview.retention.preview_id)

        assert commit.success and not commit.preview
        assert commit.retention == preview.retention
        assert commit.package_hash == preview.package_hash
        assert commit.result_version == preview.result_version == session.get_version()
        assert commit.steps[0].results[0].created[0].id == created
        assert not any("may be generated" in warning for warning in commit.warnings)
        markdown = session.project().markdown
        assert "committed as previewed" in markdown and created in markdown
        assert session.undo() and "committed as previewed" not in session.project().markdown


def test_stale_commit_refuses_without_editing(tour_plan_bytes: bytes) -> None:
    with open_session(tour_plan_bytes) as session:
        anchor = _first_paragraph(session)
        preview = session.preview_batch(_insert(anchor, "never lands"), retain=True)
        assert preview.retention is not None
        session.execute_batch(_insert(anchor, "intervening"))
        version = session.get_version()

        stale = session.commit_preview(preview.retention.preview_id)

        assert not stale.success
        assert stale.failure is not None
        assert stale.failure.error.code is EditErrorCode.PREVIEW_STALE
        assert session.get_version() == version
        assert "never lands" not in session.project().markdown


def test_a_preview_that_was_not_retained_or_failed_cannot_be_committed(
    tour_plan_bytes: bytes,
) -> None:
    with open_session(tour_plan_bytes) as session:
        anchor = _first_paragraph(session)
        plain = session.preview_batch(_insert(anchor, "plain"))
        assert plain.success and plain.retention is None

        failed = session.preview_batch(_insert("p:body:missing", "x"), retain=True)
        assert not failed.success and failed.retention is None
        assert any("not retained" in warning for warning in failed.warnings)

        missing = session.commit_preview("pv-nope")
        assert missing.failure is not None
        assert missing.failure.error.code is EditErrorCode.PREVIEW_NOT_FOUND


def test_commit_under_a_transaction_id_replays_and_is_consumed(tour_plan_bytes: bytes) -> None:
    with open_session(tour_plan_bytes) as session:
        anchor = _first_paragraph(session)
        preview = session.preview_batch(_insert(anchor, "once when retried"), retain=True)
        assert preview.retention is not None
        preview_id = preview.retention.preview_id

        first = session.commit_preview(preview_id, transaction_id="tx-commit")
        retry = session.commit_preview(preview_id, transaction_id="tx-commit")

        assert first.success and retry == first
        assert first.transaction is not None and first.transaction.transaction_id == "tx-commit"
        assert session.project().markdown.count("once when retried") == 1

        again = session.commit_preview(preview_id)
        assert again.failure is not None
        assert again.failure.error.code is EditErrorCode.PREVIEW_NOT_FOUND
