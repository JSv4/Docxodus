"""Version and optimistic-mutation transport coverage (issue #447)."""

from __future__ import annotations

from docx_scalpel import (
    EditErrorCode,
    MutationPreconditions,
    ReplaceOptions,
    TextRangePrecondition,
    open_session,
)


def test_version_and_structured_preconditions(tour_plan_bytes: bytes) -> None:
    with open_session(tour_plan_bytes) as session:
        target = next(
            a for a in session.project().anchor_index.values()
            if a.scope == "body" and a.kind in ("p", "h", "li")
        )
        info = session.get_anchor_info(target.id)
        assert info is not None
        assert info.content_hash
        assert info.visible_text is not None
        assert session.get_version() == 0

        guards = MutationPreconditions(
            expected_version=0,
            anchor_id=target.id,
            expected_content_hash=info.content_hash,
            expected_text=info.visible_text,
            expected_text_range=TextRangePrecondition(0, 0, ""),
            expected_kind=info.kind,
            expected_scope=info.scope,
        )
        assert session.check_preconditions(guards).success
        with session.preconditioned(guards):
            assert session.replace_text(target.id, "cat cat").success
        assert session.get_version() == 1

        with session.preconditioned(MutationPreconditions(expected_version=0)):
            stale = session.delete_block(target.id)
        assert not stale.success
        assert stale.error is not None
        assert stale.error.code is EditErrorCode.PRECONDITION_FAILED
        assert stale.error.precondition is not None
        assert stale.error.precondition.condition == "document_version"
        assert stale.error.precondition.expected == 0
        assert stale.error.precondition.actual == 1
        assert stale.error.precondition.current_version == 1
        assert stale.error.precondition.current_target is not None
        assert stale.error.precondition.current_target.visible_text == "cat cat"
        assert session.get_version() == 1

        count_failure = session.replace_text_range(
            target.id,
            "cat",
            "dog",
            ReplaceOptions(expected_match_count=1),
        )
        assert len(count_failure) == 1
        assert not count_failure[0].success
        assert count_failure[0].error is not None
        assert count_failure[0].error.code is EditErrorCode.PRECONDITION_FAILED
        assert count_failure[0].error.precondition is not None
        assert count_failure[0].error.precondition.condition == "match_count"
        assert session.get_version() == 1

        replaced = session.replace_text_range(
            target.id,
            "cat",
            "dog",
            ReplaceOptions(
                expected_match_count=2,
                preconditions=MutationPreconditions(expected_version=1),
            ),
        )
        assert len(replaced) == 2
        assert all(r.success for r in replaced)
        assert session.get_version() == 2
