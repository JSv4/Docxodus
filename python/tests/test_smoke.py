"""End-to-end smoke test — Python mirror of ``DocxSessionSmokeTest.cs``.

Exercises every tier of the DocxSession API on a real legal-style fixture
and confirms edits round-trip through ``save()`` and re-open. This is the
v1 acceptance gate: if this test passes, every code path the Python wrapper
exposes is hooked up to a real op on the .NET side.
"""

from __future__ import annotations

import io
import zipfile

from docx_scalpel import (
    CharSpan,
    DocxSession,
    EditResult,
    FormatOp,
    Position,
    open_session,
)


def _initial_order_index(markdown: str, anchor_id: str) -> int:
    """Same helper as the C# test — sort search results by document order."""
    needle = "{#" + anchor_id + "}"
    i = markdown.find(needle)
    return i if i >= 0 else 1 << 30


def test_ds999_agentic_workflow_on_real_document(tour_plan_bytes: bytes) -> None:
    # ── Arrange ──────────────────────────────────────────────────────────
    session = open_session(tour_plan_bytes)
    try:
        # ── Read: project to markdown ────────────────────────────────────
        initial = session.project()
        assert initial.markdown, "initial projection should not be empty"
        assert len(initial.anchor_index) > 0, "initial projection should have anchors"

        first_paragraph = min(
            (
                t
                for t in initial.anchor_index.values()
                if t.kind in ("p", "h", "li") and t.scope == "body"
            ),
            key=lambda t: _initial_order_index(initial.markdown, t.id),
        )
        anchor_id = first_paragraph.id

        # ── Mutation 1: ReplaceText (Tier A) ─────────────────────────────
        r1 = session.replace_text(
            anchor_id,
            "**SMOKETESTMARKER1:** Agentically replaced opening paragraph.",
        )
        _assert_ok(r1, "replace_text")
        assert r1.patch is not None

        # ── Mutation 2: InsertParagraph After (Tier B) ───────────────────
        r2 = session.insert_paragraph(
            anchor_id,
            Position.AFTER,
            "## Inserted Heading\n\nAgentically inserted body paragraph below the heading.",
        )
        _assert_ok(r2, "insert_paragraph")
        assert len(r2.created) == 2, f"expected 2 created anchors, got {len(r2.created)}"
        new_heading = r2.created[0]
        assert new_heading.kind == "h", f"expected heading kind=h, got {new_heading.kind}"

        # ── Mutation 3: SplitParagraph (Tier B) ──────────────────────────
        proj_for_split = session.project()
        splittable = next(
            (
                t
                for t in proj_for_split.anchor_index.values()
                if t.kind == "p"
                and t.scope == "body"
                and t.id != anchor_id
                and (info := session.get_anchor_info(t.id)) is not None
                and len(info.text_preview) > 10
            ),
            None,
        )
        if splittable is not None:
            r3 = session.split_paragraph(splittable.id, 5)
            _assert_ok(r3, "split_paragraph")
            assert len(r3.created) == 1

        # ── Mutation 4: ApplyFormat span (Tier C) ────────────────────────
        proj_for_format = session.project()
        if anchor_id in proj_for_format.anchor_index:
            r4 = session.apply_format(anchor_id, CharSpan(0, 3), FormatOp(bold=True))
            _assert_ok(r4, "apply_format")

        # ── Mutation 5: Raw escape hatch ─────────────────────────────────
        raw_xml = session.raw.get_xml(anchor_id)
        assert "w:p" in raw_xml
        modified_raw = raw_xml.replace(
            "</w:p>",
            '<w:r><w:t xml:space="preserve"> RAWINJECTED</w:t></w:r></w:p>',
        )
        r5 = session.raw.replace_xml(anchor_id, modified_raw)
        _assert_ok(r5, "raw.replace_xml")

        after_edits = session.project()
        assert "SMOKETESTMARKER1" in after_edits.markdown
        assert "Inserted Heading" in after_edits.markdown
        assert "RAWINJECTED" in after_edits.markdown

        # ── Save + reopen round-trip ─────────────────────────────────────
        saved = session.save()
        assert saved, "save() returned empty bytes"

        with open_session(saved) as reopened:
            reprojected = reopened.project()
            assert "SMOKETESTMARKER1" in reprojected.markdown
            assert "Inserted Heading" in reprojected.markdown
            assert "RAWINJECTED" in reprojected.markdown

        # ── Saved bytes are a valid DOCX (ZIP with [Content_Types].xml) ──
        _assert_is_docx(saved)

        # ── Undo all the way back ─────────────────────────────────────────
        undo_count = 0
        while session.undo():
            undo_count += 1
        final_proj = session.project()
        assert "SMOKETESTMARKER1" not in final_proj.markdown, (
            f"undo did not revert marker after {undo_count} pops"
        )
        assert "RAWINJECTED" not in final_proj.markdown
    finally:
        session.close()


def _assert_ok(result: EditResult, op_name: str) -> None:
    if not result.success:
        err = result.error
        msg = f"{op_name} failed"
        if err is not None:
            msg += f": {err.code.value} — {err.message}"
            if err.anchor_id:
                msg += f" (anchorId={err.anchor_id})"
        raise AssertionError(msg)


def _assert_is_docx(b: bytes) -> None:
    """Confirm the bytes are a syntactically-valid DOCX (a ZIP with the OOXML manifest)."""
    with zipfile.ZipFile(io.BytesIO(b)) as zf:
        names = set(zf.namelist())
    assert "[Content_Types].xml" in names, "saved bytes are not a valid OOXML package"
    assert "word/document.xml" in names, "saved bytes missing word/document.xml"


def test_ping_returns_host_metadata() -> None:
    """Sanity: the host is reachable and reports a version."""
    from docx_scalpel import ping

    pong = ping()
    assert pong.get("pong") is True
    assert isinstance(pong.get("version"), str)
