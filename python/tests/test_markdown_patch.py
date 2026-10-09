"""Block-scoped markdown patches (issue #1022) cross the wire with their blocks."""

from __future__ import annotations

from docx_scalpel import MarkdownPatch, MarkdownPatchBlock, open_session


def _top_level_blocks(session) -> list[str]:
    """Anchor ids of the body blocks the projection renders as their own lines, in order."""
    projection = session.project()
    return [
        t.id
        for t in sorted(
            (t for t in projection.anchor_index.values() if t.kind in ("p", "h") and t.scope == "body"),
            key=lambda t: projection.markdown.find("{#" + t.id + "}"),
        )
        if ("\n{#" + t.id + "}") in projection.markdown
    ]


def test_replace_text_patch_carries_only_the_edited_block(tour_plan_bytes: bytes) -> None:
    with open_session(tour_plan_bytes) as session:
        target = _top_level_blocks(session)[0]
        result = session.replace_text(target, "PATCHMARKER block-scoped replacement.")
        assert result.success
        patch = result.patch
        assert isinstance(patch, MarkdownPatch)
        assert not patch.full_document
        assert len(patch.blocks) == 1
        block = patch.blocks[0]
        assert isinstance(block, MarkdownPatchBlock)
        assert block.anchor_id == target
        assert block.after_anchor_id is None
        assert "PATCHMARKER" in block.markdown
        assert patch.markdown == block.markdown
        assert patch.removed_anchor_ids == ()


def test_patch_after_undo_covers_only_the_next_op(tour_plan_bytes: bytes) -> None:
    with open_session(tour_plan_bytes) as session:
        first, *rest = _top_level_blocks(session)
        assert session.replace_text(first, "One.").success
        assert session.undo()
        patch = session.replace_text(rest[-1], "Two.").patch
        assert patch is not None and not patch.full_document
        assert [b.anchor_id for b in patch.blocks] == [rest[-1]]
        assert "Two." in patch.markdown
