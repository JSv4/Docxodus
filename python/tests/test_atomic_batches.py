"""Atomic and explicit best-effort mutation batches (issue #445)."""

from __future__ import annotations

from docx_scalpel import (
    EditErrorCode,
    DocxSession,
    DocxSessionSettings,
    MutationBatchMode,
    MutationBatchResult,
    MutationBatchStep,
    MutationPreviewHtmlMode,
    Position,
    TableAnchorEntityKind,
    open_session,
)


def _body_paragraphs(session: DocxSession) -> list[str]:
    projection = session.project()
    return [
        anchor.id
        for anchor in projection.anchor_index.values()
        if anchor.scope == "body" and anchor.kind in ("p", "h", "li")
    ]


def test_atomic_batch_rolls_back_and_preserves_history(tour_plan_bytes: bytes) -> None:
    with open_session(tour_plan_bytes) as session:
        target = _body_paragraphs(session)[0]
        before = session.project().markdown

        result = session.execute_batch(
            [
                MutationBatchStep(
                    "replace_text",
                    {"anchorId": target, "markdown": "Speculative Python edit."},
                ),
                MutationBatchStep(
                    "replace_text",
                    {"anchorId": "p:body:missing", "markdown": "must fail"},
                ),
            ]
        )

        assert not result.success
        assert result.rolled_back
        assert result.mode is MutationBatchMode.ATOMIC
        assert result.failure is not None
        assert result.failure.index == 1
        assert result.failure.tool == "docx_scalpel"
        assert result.failure.action == "replace_text"
        assert result.failure.error.code is EditErrorCode.ANCHOR_NOT_FOUND
        assert result.failure.rolled_back
        assert all(step.rolled_back for step in result.steps)
        assert session.project().markdown == before
        assert session.get_version() == 0
        assert not session.undo()


def test_atomic_success_is_one_version_and_undo_unit(tour_plan_bytes: bytes) -> None:
    with open_session(tour_plan_bytes) as session:
        targets = _body_paragraphs(session)[:2]

        result = session.execute_batch(
            [
                MutationBatchStep(
                    "replace_text",
                    {"anchorId": targets[0], "markdown": "Python batch first."},
                ),
                MutationBatchStep(
                    "replace_text",
                    {"anchorId": targets[1], "markdown": "Python batch second."},
                ),
            ]
        )

        assert result.success
        assert result.status == "ok"
        assert session.get_version() == 1
        assert session.undo()
        assert "Python batch first." not in session.project().markdown
        assert not session.undo()


def test_best_effort_is_explicit_and_invalid_steps_are_structured(
    tour_plan_bytes: bytes,
) -> None:
    with open_session(tour_plan_bytes) as session:
        target = _body_paragraphs(session)[0]
        result = session.execute_batch(
            [
                MutationBatchStep(
                    "replace_text",
                    {"anchorId": target, "markdown": "Retained Python edit."},
                ),
                MutationBatchStep(
                    "replace_text",
                    {"anchorId": "p:body:missing", "markdown": "failure"},
                ),
            ],
            MutationBatchMode.BEST_EFFORT,
        )

        assert not result.success
        assert not result.rolled_back
        assert result.status == "partial"
        assert result.failure is not None
        assert not result.failure.rolled_back
        assert "Retained Python edit." in session.project().markdown
        assert session.get_version() == 1

    with open_session(tour_plan_bytes) as session:
        invalid = session.execute_batch([MutationBatchStep("get_version")])
        assert not invalid.success
        assert invalid.rolled_back
        assert invalid.failure is not None
        assert invalid.failure.error.code is EditErrorCode.INVALID_BATCH_STEP
        assert session.get_version() == 0


def test_preview_batch_is_rich_and_preserves_live_bytes_version_and_redo(
    tour_plan_bytes: bytes,
) -> None:
    with open_session(
        tour_plan_bytes,
        DocxSessionSettings(undo_depth=1, persist_anchor_ids=True),
    ) as session:
        targets = _body_paragraphs(session)[:2]
        assert session.replace_text(targets[0], "Python redo target.").success
        assert session.undo()
        before_version = session.get_version()
        # Match Save call sequences on each side; Save itself warms serialization caches.
        session.save(False)
        session.save(True)
        before_clean = session.save(False)
        before_persisted = session.save(True)

        result = session.preview_batch(
            [
                MutationBatchStep(
                    "replace_text",
                    {"anchorId": targets[0], "markdown": "Predicted Python first."},
                ),
                MutationBatchStep(
                    "replace_text",
                    {"anchorId": targets[1], "markdown": "Predicted Python second."},
                ),
            ],
            html_mode="scoped",
            html_anchor_id=targets[0],
        )

        assert result.preview
        assert result.success
        assert result.mode is MutationBatchMode.ATOMIC
        assert result.base_version == before_version
        assert result.result_version == before_version + 1
        assert len(result.package_hash) == 64
        assert len(result.steps) == 2
        assert result.revision_changes.added == ()
        assert result.comment_changes.added == ()
        assert result.annotation_changes.added == ()
        assert result.html is not None
        assert "Predicted Python first." in result.html

        assert session.get_version() == before_version
        assert session.save(False) == before_clean
        assert session.save(True) == before_persisted
        assert "Predicted Python" not in session.project().markdown
        assert not session.undo()
        assert session.redo()
        assert "Python redo target." in session.project().markdown


def test_absent_package_hash_decodes_to_none_not_an_empty_sentinel() -> None:
    """An unavailable hash must never satisfy a replay-equality assertion."""
    absent = MutationBatchResult._from_wire({"mode": "atomic", "packageHash": None})
    other_absent = MutationBatchResult._from_wire({"mode": "atomic"})
    assert absent.package_hash is None
    assert other_absent.package_hash is None

    present = MutationBatchResult._from_wire({"mode": "atomic", "packageHash": "ab" * 32})
    assert present.package_hash == "ab" * 32


def test_preview_html_mode_is_an_enum_matching_the_wire_strings() -> None:
    assert MutationPreviewHtmlMode.NONE.value == "none"
    assert MutationPreviewHtmlMode("scoped") is MutationPreviewHtmlMode.SCOPED
    assert MutationPreviewHtmlMode("full") is MutationPreviewHtmlMode.FULL


def test_structural_table_steps_are_batchable(tour_plan_bytes: bytes) -> None:
    """Table edits are in scope for #445, and the batch surface must match the agent one."""
    with open_session(tour_plan_bytes) as session:
        target = _body_paragraphs(session)[0]

        result = session.execute_batch(
            [
                MutationBatchStep(
                    "insert_table",
                    {
                        "anchorId": target,
                        "position": Position.AFTER.value,
                        "rows": 2,
                        "columns": 2,
                    },
                ),
                MutationBatchStep(
                    "replace_text",
                    {"anchorId": target, "markdown": "Table batch anchor."},
                ),
            ]
        )

        assert result.success, result.failure
        assert len(result.steps) == 2
        # The receipt must carry the cell-anchor map, or the caller cannot address the
        # cells the same batch just created.
        mapping = result.steps[0].results[0].table_anchors
        assert mapping is not None
        assert any(loc.entity_kind is TableAnchorEntityKind.CELL for loc in mapping.added)
        assert session.get_version() == 1
