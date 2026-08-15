#!/usr/bin/env python3
"""Generate the epic #435 acceptance workflow for `mcp_probe.py`.

The preview, the apply, and the retry must send a *byte-identical* step array —
that is what the transaction fingerprint is computed over, and a retry whose
fingerprint differs is a `transaction_conflict` rather than a replay. Writing the
array once and emitting it three times is the point of generating this file
instead of hand-maintaining the JSON.

Run:  python3 tools/mcp-server/smoke/build_epic_435_workflow.py
"""

from __future__ import annotations

import json
from pathlib import Path

TRANSACTION_ID = "epic435-smoke-restated-certificate"
CORPORATION = "Northstar Robotics, Inc."
PLACEHOLDER = "[_______________]"
RENDERER = "epic435-smoke-renderer-1"


def step(tool: str, **args: object) -> dict:
    return {"tool": tool, "args": args}


# The substantive tracked edit: a counsel filling in and annotating a restated
# certificate. Every step here is representable as a tracked revision. Bookmark
# and hyperlink mutations deliberately are not, and neither is InsertTable on this
# document shape — the engine fails those closed under render_inline rather than
# emitting markup it could not reverse — so they run in an untracked batch below.
TRACKED_STEPS = [
    step(
        "docxodus_edit",
        action="replace_text_range",
        anchorId="$name_anchor",
        find=PLACEHOLDER,
        replace=CORPORATION,
    ),
    step(
        "docxodus_create",
        action="insert_paragraph",
        anchorId="$certify_anchor",
        position="after",
        markdown="The undersigned further certifies that this Restated Certificate "
        "has been duly adopted in accordance with Sections 242 and 245 of the DGCL.",
    ),
    step(
        "docxodus_create",
        action="insert_footnote",
        anchorId="$certify_anchor",
        characterOffset=19,
        markdown="Adopted by written consent of the stockholders.",
    ),
    step(
        "docxodus_comment",
        action="add",
        anchorId="$name_anchor",
        author="Smoke Counsel",
        initials="SC",
        markdown="Confirm the exact legal name against the charter before filing.",
    ),
    # Bolds the name this batch's first step just inserted. Addressing text an
    # earlier step inserted in the SAME tracked batch is exactly what DS409 fixed.
    step(
        "docxodus_format",
        action="apply_format_by_substring",
        anchorId="$name_anchor",
        substring=CORPORATION,
        format={"bold": True},
    ),
]


def mutations(call_id: str, mode: str, **extra: object) -> dict:
    arguments: dict = {"sessionId": "$session_id", "mode": mode}
    arguments.update(extra.pop("arguments_extra", {}))  # type: ignore[arg-type]
    arguments["preconditions"] = {"expectedVersion": "$base_version"}
    arguments["steps"] = TRACKED_STEPS
    call: dict = {"id": call_id, "name": "docxodus_mutations", "arguments": arguments}
    call.update(extra)
    return call


def workflow() -> list[dict]:
    calls: list[dict] = []

    # ── Inspect ──────────────────────────────────────────────────────────
    calls += [
        {
            "id": "open",
            "name": "docxodus_open",
            "arguments": {
                "path": "local.docx",
                "trackedChanges": "render_inline",
                "revisionAuthor": "Smoke Counsel",
                "persistAnchorIds": True,
            },
            "capture": {"session_id": "sessionId"},
        },
        {
            "id": "inspect_info",
            "name": "docxodus_get_content",
            "arguments": {"sessionId": "$session_id", "format": "info"},
            "capture": {
                "base_version": "version",
                "total_anchors": "editSummary.totalAnchors",
                "footnote_count": "editSummary.footnoteCount",
            },
            "expect": {"version": 0, "sectionInfo.pageNumberFormat": "lowerRoman"},
        },
        {
            "id": "inspect_name_clause",
            "name": "docxodus_search",
            "arguments": {
                "sessionId": "$session_id",
                "mode": "text",
                "query": "The name of this corporation is",
                "maxResults": 2,
            },
            "capture": {"name_anchor": "matches.1.enclosingAnchor.id"},
        },
        {
            "id": "inspect_certification",
            "name": "docxodus_search",
            "arguments": {
                "sessionId": "$session_id",
                "mode": "text",
                "query": "DOES HEREBY CERTIFY",
                "maxResults": 1,
            },
            "capture": {"certify_anchor": "matches.0.enclosingAnchor.id"},
        },
        {
            "id": "inspect_list_item",
            "name": "docxodus_search",
            "arguments": {
                "sessionId": "$session_id",
                "mode": "kind",
                "query": "li",
                "maxResults": 1,
            },
            "capture": {"list_anchor": "matches.0.id"},
        },
        {
            "id": "inspect_bookmarks",
            "name": "docxodus_links",
            "arguments": {
                "sessionId": "$session_id",
                "action": "list_bookmarks",
                "scope": "body",
            },
            "capture": {"bookmarks_before": "bookmarks.length"},
        },
        {
            "id": "inspect_revisions",
            "name": "docxodus_track_changes",
            "arguments": {"sessionId": "$session_id", "action": "list"},
            "expect": {"revisions.length": 0},
        },
        {
            "id": "inspect_pagination_unregistered",
            "name": "docxodus_pagination",
            "arguments": {"sessionId": "$session_id", "action": "status"},
            "expect": {
                "availability": "unavailable",
                "unavailableReason": "no_page_map",
            },
        },
    ]

    # ── Preview, and prove the live session did not move ──────────────────
    calls += [
        mutations(
            "preview",
            "preview",
            arguments_extra={"previewHtml": "full"},
            capture={
                "predicted_version": "resultVersion",
                "predicted_hash": "packageHash",
                "predicted_revision_adds": "revisionChanges.added.length",
                "predicted_comment_adds": "commentChanges.added.length",
                # Affected anchors, per step. Counts for created anchors, not ids: the
                # receipt warns that generated ids differ between preview and apply.
                "predicted_paragraph_creates": "steps.1.results.0.created.length",
                "predicted_footnote_creates": "steps.2.results.0.created.length",
                "predicted_comment_creates": "steps.3.results.0.created.length",
            },
            expect={
                "status": "ok",
                "preview": True,
                "rolledBack": False,
                "baseVersion": "$base_version",
                "editsApplied": len(TRACKED_STEPS),
                "steps.0.results.0.modified.0.id": "$name_anchor",
                "steps.2.results.0.modified.0.id": "$certify_anchor",
                "steps.4.results.0.modified.0.id": "$name_anchor",
                "steps.0.results.0.removed.length": 0,
                "steps.1.results.0.removed.length": 0,
            },
        ),
        {
            "id": "preview_left_live_untouched",
            "name": "docxodus_get_content",
            "arguments": {"sessionId": "$session_id", "format": "info"},
            "expect": {
                "version": "$base_version",
                "editSummary.totalAnchors": "$total_anchors",
                "editSummary.footnoteCount": "$footnote_count",
            },
        },
        {
            "id": "preview_authored_no_live_revisions",
            "name": "docxodus_track_changes",
            "arguments": {"sessionId": "$session_id", "action": "list"},
            "expect": {"revisions.length": 0},
        },
        {
            "id": "preview_authored_no_live_comments",
            "name": "docxodus_comment",
            "arguments": {"sessionId": "$session_id", "action": "list"},
            "expect": {"comments.length": 0},
        },
    ]

    # ── Apply atomically under a transaction id ───────────────────────────
    calls += [
        mutations(
            "apply",
            "atomic",
            arguments_extra={"transactionId": TRANSACTION_ID},
            capture={"applied_version": "resultVersion", "applied_hash": "packageHash"},
            expect={
                "status": "ok",
                "preview": False,
                "rolledBack": False,
                "baseVersion": "$base_version",
                "editsApplied": len(TRACKED_STEPS),
                # The preview's predictions, now checked against reality.
                "resultVersion": "$predicted_version",
                "revisionChanges.added.length": "$predicted_revision_adds",
                "commentChanges.added.length": "$predicted_comment_adds",
                # The anchors the preview said would be touched are the ones that were.
                "steps.0.results.0.modified.0.id": "$name_anchor",
                "steps.2.results.0.modified.0.id": "$certify_anchor",
                "steps.4.results.0.modified.0.id": "$name_anchor",
                "steps.0.results.0.removed.length": 0,
                "steps.1.results.0.removed.length": 0,
                "steps.1.results.0.created.length": "$predicted_paragraph_creates",
                "steps.2.results.0.created.length": "$predicted_footnote_creates",
                "steps.3.results.0.created.length": "$predicted_comment_creates",
                "transaction.transactionId": TRANSACTION_ID,
            },
        ),
        {
            "id": "apply_advanced_live_state",
            "name": "docxodus_get_content",
            "arguments": {"sessionId": "$session_id", "format": "info"},
            "capture": {
                "live_version": "version",
                "live_anchors": "editSummary.totalAnchors",
                "live_footnotes": "editSummary.footnoteCount",
            },
            "expect": {"version": "$applied_version"},
        },
        {
            "id": "apply_authored_revisions",
            "name": "docxodus_track_changes",
            "arguments": {"sessionId": "$session_id", "action": "list"},
            "capture": {"revisions_after_apply": "revisions.length"},
        },
        {
            "id": "apply_authored_comment",
            "name": "docxodus_comment",
            "arguments": {"sessionId": "$session_id", "action": "list"},
            "expect": {"comments.length": 1, "comments.0.author": "Smoke Counsel"},
        },
    ]

    # ── Simulated response loss: retry the identical request ──────────────
    calls += [
        mutations(
            "retry_after_simulated_response_loss",
            "atomic",
            arguments_extra={"transactionId": TRANSACTION_ID},
            expectSameAs="apply",
        ),
        {
            "id": "retry_created_no_duplicate_edit",
            "name": "docxodus_get_content",
            "arguments": {"sessionId": "$session_id", "format": "info"},
            "expect": {
                "version": "$live_version",
                "editSummary.totalAnchors": "$live_anchors",
                "editSummary.footnoteCount": "$live_footnotes",
            },
        },
        {
            "id": "retry_created_no_duplicate_revisions",
            "name": "docxodus_track_changes",
            "arguments": {"sessionId": "$session_id", "action": "list"},
            "expect": {"revisions.length": "$revisions_after_apply"},
        },
        {
            "id": "retry_created_no_duplicate_comment",
            "name": "docxodus_comment",
            "arguments": {"sessionId": "$session_id", "action": "list"},
            "expect": {"comments.length": 1},
        },
    ]

    # ── Link scaffolding: the engine fails closed under tracked recording ──
    calls += [
        {
            "id": "bookmark_is_refused_while_tracking",
            "name": "docxodus_links",
            "arguments": {
                "sessionId": "$session_id",
                "action": "add_bookmark",
                "name": "SmokeCorporationName",
                "startAnchorId": "$name_anchor",
                "startOffset": 2,
                "endAnchorId": "$name_anchor",
                "endOffset": 26,
            },
            "expectFailure": True,
            "expect": {"error.code": "tracked_operation_unsupported"},
        },
        {
            "id": "table_insert_is_refused_while_tracking",
            "name": "docxodus_create",
            "arguments": {
                "sessionId": "$session_id",
                "action": "insert_table",
                "anchorId": "$certify_anchor",
                "position": "after",
                "rows": 2,
                "columns": 2,
                "cellContents": ["Class", "Authorized Shares"],
            },
            "expectFailure": True,
            "expect": {"error.code": "tracked_operation_unsupported"},
        },
        {
            "id": "switch_to_direct_recording",
            "name": "docxodus_track_changes",
            "arguments": {
                "sessionId": "$session_id",
                "action": "set_mode",
                "mode": "accept",
            },
        },
        {
            "id": "link_scaffolding_batch",
            "name": "docxodus_mutations",
            "arguments": {
                "sessionId": "$session_id",
                "mode": "atomic",
                "steps": [
                    # Deliberately NOT the name clause: its offsets now fall inside
                    # the w:ins the tracked batch authored, and a bookmark endpoint
                    # inside a revision is refused (unsupported_inline_boundary).
                    step(
                        "docxodus_links",
                        action="add_bookmark",
                        name="SmokeDelawareRationale",
                        startAnchorId="$list_anchor",
                        startOffset=12,
                        endAnchorId="$list_anchor",
                        endOffset=30,
                    ),
                    step(
                        "docxodus_links",
                        action="add_hyperlink",
                        anchorId="$list_anchor",
                        startOffset=0,
                        length=8,
                        kind="external",
                        target="https://delcode.delaware.gov/title8/c001/",
                    ),
                    step(
                        "docxodus_create",
                        action="insert_table",
                        anchorId="$certify_anchor",
                        position="after",
                        rows=2,
                        columns=2,
                        cellContents=[
                            "Class",
                            "Authorized Shares",
                            "Common Stock",
                            "100,000,000",
                        ],
                    ),
                ],
            },
            "expect": {"status": "ok", "editsApplied": 3},
        },
        {
            "id": "table_metadata_is_addressable",
            "name": "docxodus_search",
            "arguments": {
                "sessionId": "$session_id",
                "mode": "kind",
                "query": "tbl",
                "maxResults": 1,
            },
            "capture": {"table_anchor": "matches.0.id"},
        },
        {
            "id": "table_coordinates_resolve",
            "name": "docxodus_table",
            "arguments": {
                "sessionId": "$session_id",
                "action": "resolve_cell_coordinate",
                "tableAnchorId": "$table_anchor",
                "rowIndex": 1,
                "columnIndex": 1,
            },
        },
        {
            "id": "bookmark_resolves_to_its_anchor",
            "name": "docxodus_search",
            "arguments": {
                "sessionId": "$session_id",
                "mode": "bookmark",
                "query": "SmokeDelawareRationale",
            },
            "expect": {"matches.0.id": "$list_anchor"},
        },
        {
            "id": "hyperlink_is_registered",
            "name": "docxodus_links",
            "arguments": {
                "sessionId": "$session_id",
                "action": "list_hyperlinks",
                "scope": "body",
            },
            "expect": {"hyperlinks.length": 1, "hyperlinks.0.kind": "external"},
        },
        {
            "id": "resume_tracked_recording",
            "name": "docxodus_track_changes",
            "arguments": {
                "sessionId": "$session_id",
                "action": "set_mode",
                "mode": "render_inline",
                "revisionAuthor": "Smoke Counsel",
            },
        },
        {
            "id": "state_after_link_scaffolding",
            "name": "docxodus_get_content",
            "arguments": {"sessionId": "$session_id", "format": "info"},
            "capture": {
                "guarded_version": "version",
                "guarded_anchors": "editSummary.totalAnchors",
            },
        },
        {
            "id": "revisions_after_link_scaffolding",
            "name": "docxodus_track_changes",
            "arguments": {"sessionId": "$session_id", "action": "list"},
            "capture": {"guarded_revisions": "revisions.length"},
        },
    ]

    # ── Intentional stale request and intentional atomic failure ──────────
    calls += [
        {
            "id": "stale_precondition_is_refused",
            "name": "docxodus_edit",
            "arguments": {
                "sessionId": "$session_id",
                "action": "replace_text_range",
                "anchorId": "$name_anchor",
                "find": "The name of this corporation is",
                "replace": "This must never reach the document.",
                "preconditions": {"expectedVersion": "$base_version"},
            },
            "expectFailure": True,
            "expect": {
                "error.code": "precondition_failed",
                "error.precondition.currentVersion": "$guarded_version",
            },
        },
        {
            "id": "atomic_batch_failure_rolls_back",
            "name": "docxodus_mutations",
            "arguments": {
                "sessionId": "$session_id",
                "mode": "atomic",
                "steps": [
                    step(
                        "docxodus_create",
                        action="insert_paragraph",
                        anchorId="$name_anchor",
                        position="after",
                        markdown="A first edit that must be rolled back.",
                    ),
                    step(
                        "docxodus_edit",
                        action="replace_text_range",
                        anchorId="$name_anchor",
                        find=CORPORATION,
                        replace="A second edit that must be rolled back.",
                    ),
                    step(
                        "docxodus_edit",
                        action="replace_text",
                        anchorId="p:body:0000000000000000deadbeefdeadbeef",
                        markdown="This step fails on a nonexistent anchor.",
                    ),
                ],
            },
            "expectFailure": True,
            "expect": {
                "status": "failed",
                "rolledBack": True,
                "editsApplied": 0,
                "failure.index": 2,
                "failure.error.code": "anchor_not_found",
            },
        },
    ]

    # ── Audit: content, version, history, anchors and revisions held ──────
    calls += [
        {
            "id": "audit_version_and_anchors_unchanged",
            "name": "docxodus_get_content",
            "arguments": {"sessionId": "$session_id", "format": "info"},
            "expect": {
                "version": "$guarded_version",
                "editSummary.totalAnchors": "$guarded_anchors",
            },
        },
        {
            "id": "audit_revisions_unchanged",
            "name": "docxodus_track_changes",
            "arguments": {"sessionId": "$session_id", "action": "list"},
            "expect": {"revisions.length": "$guarded_revisions"},
        },
        {
            "id": "audit_rolled_back_text_absent",
            "name": "docxodus_search",
            "arguments": {
                "sessionId": "$session_id",
                "mode": "text",
                "query": "must be rolled back",
                "scope": "all",
            },
            "expect": {"matches.length": 0},
        },
        {
            # The applied edit is visible as a revision with the exact inserted text.
            "id": "audit_applied_text_present_in_revisions",
            "name": "docxodus_track_changes",
            "arguments": {"sessionId": "$session_id", "action": "list"},
            "expect": {
                "revisions.length": 4,
                "revisions.1.type": "delete",
                "revisions.1.text": PLACEHOLDER,
                "revisions.2.type": "insert",
                "revisions.2.text": CORPORATION,
                "revisions.2.author": "Smoke Counsel",
                "revisions.3.type": "format",
                "revisions.3.text": CORPORATION,
            },
        },
        {
            # Regression guard for the defect this smoke found: text inserted under
            # render_inline lives in w:ins, which InlineRuns used to skip, so search
            # could not see the edit it had just made (fixed with DS409/DS410).
            "id": "applied_text_is_findable_after_tracked_edit",
            "name": "docxodus_search",
            "arguments": {
                "sessionId": "$session_id",
                "mode": "text",
                "query": CORPORATION,
                "scope": "body",
            },
            "expect": {"matches.0.enclosingAnchor.id": "$name_anchor"},
        },
        {
            "id": "audit_undo_redo_still_usable",
            "name": "docxodus_edit",
            "arguments": {"sessionId": "$session_id", "action": "undo"},
            "expect": {"success": True},
        },
        {
            "id": "audit_redo_restores",
            "name": "docxodus_edit",
            "arguments": {"sessionId": "$session_id", "action": "redo"},
            "expect": {"success": True},
        },
        {
            "id": "audit_state_after_undo_redo",
            "name": "docxodus_get_content",
            "arguments": {"sessionId": "$session_id", "format": "info"},
            "expect": {"editSummary.totalAnchors": "$guarded_anchors"},
            "capture": {"final_version": "version"},
        },
    ]

    # ── Page-aware citation ───────────────────────────────────────────────
    calls += [
        {
            "id": "register_page_map",
            "name": "docxodus_pagination",
            "arguments": {
                "sessionId": "$session_id",
                "action": "register",
                "pageMap": {
                    "schemaVersion": 1,
                    "mode": "paginated",
                    "availability": "available",
                    "documentVersion": "$final_version",
                    "rendererFingerprint": RENDERER,
                    "pages": [
                        {
                            "pageNumber": 1,
                            "pageInSection": 1,
                            "width": 816,
                            "height": 1056,
                            "sectionIndex": 0,
                            "pageName": "i",
                        }
                    ],
                    "fragments": [
                        {
                            "fragmentId": "f-name-0",
                            "anchorId": "$name_anchor",
                            "fragmentIndex": 0,
                            "pageNumber": 1,
                            "geometry": {
                                "x": 96,
                                "y": 240,
                                "width": 624,
                                "height": 32,
                            },
                            "story": "body",
                            "inTableCell": False,
                        }
                    ],
                },
            },
        },
        {
            "id": "cite_page_for_anchor",
            "name": "docxodus_pagination",
            "arguments": {
                "sessionId": "$session_id",
                "action": "cite",
                "anchorId": "$name_anchor",
                "citation": {
                    "documentVersion": "$final_version",
                    "rendererFingerprint": RENDERER,
                },
            },
            "expect": {
                "availability": "available",
                "documentVersion": "$final_version",
                "pages.0.pageNumber": 1,
                "pages.0.pageName": "i",
                "fragments.0.anchorId": "$name_anchor",
                "fragments.0.story": "body",
            },
        },
    ]

    # ── Persist ───────────────────────────────────────────────────────────
    calls += [
        {
            "id": "save",
            "name": "docxodus_save",
            "arguments": {
                "sessionId": "$session_id",
                "path": "smoke-output.docx",
                "persistAnchorIds": True,
            },
        },
        {"id": "close", "name": "docxodus_close", "arguments": {"sessionId": "$session_id"}},
    ]

    return calls


def main() -> None:
    target = Path(__file__).with_name("epic-435-workflow.json")
    target.write_text(json.dumps(workflow(), indent=2) + "\n", encoding="utf-8")
    print(f"wrote {target} ({len(workflow())} calls)")


if __name__ == "__main__":
    main()
