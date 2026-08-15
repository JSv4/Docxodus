"""Atomic and explicit best-effort mutation batches (issue #445)."""

from __future__ import annotations

from docx_scalpel import (
    EditErrorCode,
    DocxSession,
    MutationBatchMode,
    MutationBatchStep,
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
