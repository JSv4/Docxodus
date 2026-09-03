#!/usr/bin/env python3
"""Generate the epic #435 acceptance fixtures for `mcp_probe.py`.

Two files come out of here, and both are generated for the same reason: a
contract stated twice drifts.

`epic-435-workflow.json` — the preview, the apply, and the retry must send a
*byte-identical* step array. That is what the transaction fingerprint is
computed over, and a retry whose fingerprint differs is a
`transaction_conflict` rather than a replay. Writing the array once and
emitting it three times is the point of generating the file.

`epic-435-validation.json` — reopens the saved package and re-asserts what
persisted, including the same revision list the workflow audits live. That list
was maintained by hand in both places and went stale in both (#687), so it now
has a single declaration, `REVISION_MEMBERS`.

Run:  python3 tools/mcp-server/smoke/build_epic_435_fixtures.py
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

TRANSACTION_ID = "epic435-smoke-restated-certificate"
AUTHOR = "Smoke Counsel"
CORPORATION = "Northstar Robotics, Inc."
PLACEHOLDER = "[_______________]"
RENDERER = "epic435-smoke-renderer-1"
HISTORY_SENTINEL = "History cursor sentinel — this paragraph exists only on redo."
CERTIFICATION = (
    "The undersigned further certifies that this Restated Certificate "
    "has been duly adopted in accordance with Sections 242 and 245 of the DGCL."
)
FOOTNOTE = "Adopted by written consent of the stockholders."

BODY = "/word/document.xml"
NOTES = "/word/footnotes.xml"


def step(tool: str, **args: object) -> dict:
    return {"tool": tool, "args": args}


def revision(text: str, kind: str = "insert", **fields: object) -> dict:
    return {"type": kind, "text": text, "author": AUTHOR, **fields}


# The five tracked steps below settle into exactly six revisions, and *which* six
# is the contract this smoke guards. Two of them come from the single
# `insert_footnote` step: the reference run in the body — a genuine tracked
# insertion that carries no text, because `w:footnoteReference` has none to
# carry — and the note body over in word/footnotes.xml. The engine's enumeration
# ORDER is not part of the contract, so these are matched as members rather than
# by index (#687); `expectMembers` still demands exactly one entry per member,
# so a duplicated or vanished revision fails just as loudly.
REVISION_MEMBERS = [
    revision("", partUri=BODY, scope="body", anchorId="$certify_anchor"),
    revision(CERTIFICATION + "¶", partUri=BODY, scope="body"),
    revision(PLACEHOLDER, "delete", partUri=BODY, scope="body", anchorId="$name_anchor"),
    revision(CORPORATION, partUri=BODY, scope="body", anchorId="$name_anchor"),
    revision(CORPORATION, "format", partUri=BODY, scope="body", anchorId="$name_anchor"),
    revision(" " + FOOTNOTE + "¶", partUri=NOTES, scope="fn"),
]

# Both runs discover the two anchors REVISION_MEMBERS keys on the same way, from
# text that neither the edit script nor the save disturbs.
ANCHOR_DISCOVERY = [
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
]


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
        markdown=CERTIFICATION,
    ),
    step(
        "docxodus_create",
        action="insert_footnote",
        anchorId="$certify_anchor",
        characterOffset=19,
        markdown=FOOTNOTE,
    ),
    step(
        "docxodus_comment",
        action="add",
        anchorId="$name_anchor",
        author=AUTHOR,
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


def affected_anchor_expectations() -> dict[str, object]:
    """Exact semantic affected-anchor shape shared by preview and apply.

    Generated ids intentionally differ between executions, but cardinality, kind, scope, and
    every caller-known modified id are deterministic and are all part of the receipt contract.
    """
    expected: dict[str, object] = {}
    for index in range(len(TRACKED_STEPS)):
        expected[f"steps.{index}.success"] = True
        expected[f"steps.{index}.rolledBack"] = False
        expected[f"steps.{index}.results.length"] = 1
        expected[f"steps.{index}.results.0.success"] = True
        expected[f"steps.{index}.results.0.removed.length"] = 0

    expected.update({
        "steps.0.results.0.created.length": 0,
        "steps.0.results.0.modified.length": 1,
        "steps.0.results.0.modified.0.id": "$name_anchor",

        "steps.1.results.0.created.length": 1,
        "steps.1.results.0.created.0.kind": "p",
        "steps.1.results.0.created.0.scope": "body",
        "steps.1.results.0.modified.length": 0,

        "steps.2.results.0.created.length": 2,
        "steps.2.results.0.created.0.kind": "fn",
        "steps.2.results.0.created.0.scope": "fn",
        "steps.2.results.0.created.1.kind": "p",
        "steps.2.results.0.created.1.scope": "fn",
        "steps.2.results.0.modified.length": 1,
        "steps.2.results.0.modified.0.id": "$certify_anchor",

        "steps.3.results.0.created.length": 2,
        "steps.3.results.0.created.0.kind": "cmt",
        "steps.3.results.0.created.0.scope": "cmt",
        "steps.3.results.0.created.1.kind": "p",
        "steps.3.results.0.created.1.scope": "cmt",
        "steps.3.results.0.modified.length": 1,
        "steps.3.results.0.modified.0.id": "$name_anchor",

        "steps.4.results.0.created.length": 0,
        "steps.4.results.0.modified.length": 1,
        "steps.4.results.0.modified.0.id": "$name_anchor",
    })
    return expected


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
        *ANCHOR_DISCOVERY,
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
                "preview_html_bytes": "html.length",
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
                **affected_anchor_expectations(),
            },
            expectNonEmpty=["html", "packageHash"],
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
                # Generated ids differ, but the complete semantic affected-anchor shape must not.
                **affected_anchor_expectations(),
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
            "id": "bookmark_revision_interior_is_refused",
            "name": "docxodus_links",
            "arguments": {
                "sessionId": "$session_id",
                "action": "add_bookmark",
                "name": "SmokeRevisionInteriorShouldFail",
                # The inserted corporate name begins at visible offset 32. Both endpoints are
                # strictly inside its w:ins, so direct recording reaches the boundary guard rather
                # than the tracked-operation guard exercised above.
                "startAnchorId": "$name_anchor",
                "startOffset": 33,
                "endAnchorId": "$name_anchor",
                "endOffset": 40,
            },
            "expectFailure": True,
            "expect": {"error.code": "unsupported_inline_boundary"},
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
            # Seed a known pre-existing redo cursor without changing the guarded package. A failed
            # atomic batch must preserve the snapshot itself, not merely leave redo callable.
            "id": "seed_history_redo_target",
            "name": "docxodus_create",
            "arguments": {
                "sessionId": "$session_id",
                "action": "insert_paragraph",
                "anchorId": "$certify_anchor",
                "position": "after",
                "markdown": HISTORY_SENTINEL,
            },
            "capture": {"history_seed_anchor": "created.0.id"},
            "expect": {
                "success": True,
                "created.length": 1,
                "created.0.kind": "p",
                "created.0.scope": "body",
                "removed.length": 0,
            },
        },
        {
            "id": "seed_history_redo_cursor",
            "name": "docxodus_edit",
            "arguments": {"sessionId": "$session_id", "action": "undo"},
            "expect": {"success": True},
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
            "capture": {
                "guarded_revisions": "revisions",
                "guarded_revision_count": "revisions.length",
            },
        },
        {
            "id": "comments_before_failure",
            "name": "docxodus_comment",
            "arguments": {"sessionId": "$session_id", "action": "list"},
            "capture": {"guarded_comments": "comments"},
        },
        {
            "id": "content_and_anchor_identity_before_failure",
            "name": "docxodus_get_content",
            "arguments": {"sessionId": "$session_id", "format": "markdown"},
            "capture": {
                "guarded_markdown": "markdown",
                "guarded_anchor_index": "anchorIndex",
            },
        },
        {
            "id": "package_checkpoint_before_failure",
            "name": "docxodus_mutations",
            "arguments": {"sessionId": "$session_id", "mode": "preview", "steps": []},
            "capture": {"guarded_hash": "packageHash"},
            "expect": {
                "status": "ok",
                "preview": True,
                "editsApplied": 0,
                "baseVersion": "$guarded_version",
                "resultVersion": "$guarded_version",
            },
            "expectNonEmpty": ["packageHash"],
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
                "baseVersion": "$guarded_version",
                "resultVersion": "$guarded_version",
                "packageHash": "$guarded_hash",
                "revisionChanges.added.length": 0,
                "revisionChanges.removed.length": 0,
                "revisionChanges.modified.length": 0,
                "commentChanges.added.length": 0,
                "commentChanges.removed.length": 0,
                "commentChanges.modified.length": 0,
                "annotationChanges.added.length": 0,
                "annotationChanges.removed.length": 0,
                "annotationChanges.modified.length": 0,
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
            "id": "audit_content_and_anchor_identity_unchanged",
            "name": "docxodus_get_content",
            "arguments": {"sessionId": "$session_id", "format": "markdown"},
            "expect": {
                "markdown": "$guarded_markdown",
                "anchorIndex": "$guarded_anchor_index",
            },
        },
        {
            "id": "audit_revisions_unchanged",
            "name": "docxodus_track_changes",
            "arguments": {"sessionId": "$session_id", "action": "list"},
            "expect": {"revisions": "$guarded_revisions"},
        },
        {
            "id": "audit_comments_unchanged",
            "name": "docxodus_comment",
            "arguments": {"sessionId": "$session_id", "action": "list"},
            "expect": {"comments": "$guarded_comments"},
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
            "expect": {"revisions.length": len(REVISION_MEMBERS)},
            "expectMembers": {"revisions": REVISION_MEMBERS},
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
            "id": "audit_preexisting_redo_restores_seed",
            "name": "docxodus_edit",
            "arguments": {"sessionId": "$session_id", "action": "redo"},
            "expect": {"success": True},
        },
        {
            "id": "audit_redo_restored_exact_snapshot",
            "name": "docxodus_search",
            "arguments": {
                "sessionId": "$session_id",
                "mode": "text",
                "query": HISTORY_SENTINEL,
                "scope": "body",
            },
            "expect": {
                "matches.length": 1,
                "matches.0.enclosingAnchor.id": "$history_seed_anchor",
            },
        },
        {
            "id": "audit_undo_rewinds_seed",
            "name": "docxodus_edit",
            "arguments": {"sessionId": "$session_id", "action": "undo"},
            "expect": {"success": True},
        },
        {
            "id": "audit_package_after_redo_round_trip",
            "name": "docxodus_mutations",
            "arguments": {"sessionId": "$session_id", "mode": "preview", "steps": []},
            "expect": {"status": "ok", "packageHash": "$guarded_hash"},
        },
        {
            "id": "audit_preexisting_undo_rewinds_link_batch",
            "name": "docxodus_edit",
            "arguments": {"sessionId": "$session_id", "action": "undo"},
            "expect": {"success": True},
        },
        {
            "id": "audit_undo_removed_link_batch",
            "name": "docxodus_search",
            "arguments": {
                "sessionId": "$session_id",
                "mode": "bookmark",
                "query": "SmokeDelawareRationale",
            },
            "expect": {"matches.length": 0},
        },
        {
            "id": "audit_preexisting_undo_can_redo",
            "name": "docxodus_edit",
            "arguments": {"sessionId": "$session_id", "action": "redo"},
            "expect": {"success": True},
        },
        {
            "id": "audit_package_after_undo_round_trip",
            "name": "docxodus_mutations",
            "arguments": {"sessionId": "$session_id", "mode": "preview", "steps": []},
            "expect": {"status": "ok", "packageHash": "$guarded_hash"},
        },
        {
            "id": "audit_link_batch_restored",
            "name": "docxodus_search",
            "arguments": {
                "sessionId": "$session_id",
                "mode": "bookmark",
                "query": "SmokeDelawareRationale",
            },
            "expect": {"matches.0.id": "$list_anchor"},
        },
        {
            "id": "audit_redo_cursor_survived_undo_round_trip",
            "name": "docxodus_edit",
            "arguments": {"sessionId": "$session_id", "action": "redo"},
            "expect": {"success": True},
        },
        {
            "id": "audit_surviving_redo_restored_exact_snapshot",
            "name": "docxodus_search",
            "arguments": {
                "sessionId": "$session_id",
                "mode": "text",
                "query": HISTORY_SENTINEL,
                "scope": "body",
            },
            "expect": {
                "matches.length": 1,
                "matches.0.enclosingAnchor.id": "$history_seed_anchor",
            },
        },
        {
            "id": "audit_rewind_surviving_redo",
            "name": "docxodus_edit",
            "arguments": {"sessionId": "$session_id", "action": "undo"},
            "expect": {"success": True},
        },
        {
            "id": "audit_package_after_all_history_round_trips",
            "name": "docxodus_mutations",
            "arguments": {"sessionId": "$session_id", "mode": "preview", "steps": []},
            "expect": {"status": "ok", "packageHash": "$guarded_hash"},
        },
        {
            "id": "audit_state_after_history_round_trips",
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


def validation() -> list[dict]:
    """Reopen the saved package and assert what survived the round trip.

    Independent of the workflow's session: this run reopens `smoke-output.docx`
    from scratch and re-discovers its anchors, so what it proves is that the
    edits are in the *package*, not in a live session's memory.
    """
    return [
        {
            "id": "reopen_saved_output",
            "name": "docxodus_open",
            "arguments": {"path": "smoke-output.docx", "trackedChanges": "accept"},
            "capture": {"session_id": "sessionId"},
        },
        {
            "id": "package_reopens_with_intact_structure",
            "name": "docxodus_get_content",
            "arguments": {"sessionId": "$session_id", "format": "info"},
            "capture": {"reopened_anchors": "editSummary.totalAnchors"},
            "expect": {
                "version": 0,
                "sectionInfo.pageNumberFormat": "lowerRoman",
                "sectionInfo.headerRefs.0.kind": "first",
                "sectionInfo.footerRefs.0.kind": "even",
            },
        },
        *ANCHOR_DISCOVERY,
        {
            "id": "tracked_edits_persisted_as_revisions",
            "name": "docxodus_track_changes",
            "arguments": {"sessionId": "$session_id", "action": "list"},
            "expect": {"revisions.length": len(REVISION_MEMBERS)},
            "expectMembers": {"revisions": REVISION_MEMBERS},
        },
        {
            "id": "comment_persisted",
            "name": "docxodus_comment",
            "arguments": {"sessionId": "$session_id", "action": "list"},
            "expect": {"comments.length": 1, "comments.0.author": AUTHOR},
        },
        {
            "id": "bookmark_persisted",
            "name": "docxodus_search",
            "arguments": {
                "sessionId": "$session_id",
                "mode": "bookmark",
                "query": "SmokeDelawareRationale",
            },
            "expect": {"matches.length": 1},
        },
        {
            "id": "hyperlink_persisted_as_external_relationship",
            "name": "docxodus_links",
            "arguments": {
                "sessionId": "$session_id",
                "action": "list_hyperlinks",
                "scope": "body",
            },
            "expect": {
                "hyperlinks.length": 1,
                "hyperlinks.0.kind": "external",
                "hyperlinks.0.target": "https://delcode.delaware.gov/title8/c001/",
            },
        },
        {
            "id": "table_persisted",
            "name": "docxodus_search",
            "arguments": {"sessionId": "$session_id", "mode": "kind", "query": "tbl"},
            "expect": {"matches.length": 1},
            "capture": {"table_anchor": "matches.0.id"},
        },
        {
            "id": "table_grid_is_addressable",
            "name": "docxodus_table",
            "arguments": {
                "sessionId": "$session_id",
                "action": "get_metadata",
                "tableAnchorId": "$table_anchor",
            },
            "expect": {"metadata.rows.length": 2, "metadata.columns.length": 2},
        },
        {
            "id": "footnote_persisted",
            "name": "docxodus_search",
            "arguments": {
                "sessionId": "$session_id",
                "mode": "text",
                "query": "Adopted by written consent",
                "scope": "all",
            },
            "expect": {"matches.length": 1},
        },
        {
            "id": "inserted_paragraph_persisted_and_findable",
            "name": "docxodus_search",
            "arguments": {
                "sessionId": "$session_id",
                "mode": "text",
                "query": "The undersigned further certifies",
                "scope": "body",
            },
            "expect": {"matches.length": 1},
        },
        {
            "id": "rolled_back_text_never_persisted",
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
            "id": "refused_text_never_persisted",
            "name": "docxodus_search",
            "arguments": {
                "sessionId": "$session_id",
                "mode": "text",
                "query": "must never reach the document",
                "scope": "all",
            },
            "expect": {"matches.length": 0},
        },
        {
            "id": "close",
            "name": "docxodus_close",
            "arguments": {"sessionId": "$session_id"},
        },
    ]


def main() -> int:
    check = "--check" in sys.argv[1:]
    stale = []
    for name, calls in (
            ("epic-435-workflow.json", workflow()),
            ("epic-435-validation.json", validation())):
        target = Path(__file__).with_name(name)
        rendered = json.dumps(calls, indent=2) + "\n"
        if check:
            # Asserts the committed fixture is what this generator emits, without
            # consulting git — a working tree mid-edit is not the question.
            current = target.read_text(encoding="utf-8") if target.exists() else None
            print(f"{'stale' if current != rendered else 'fresh'}: {target}")
            if current != rendered:
                stale.append(target)
            continue
        target.write_text(rendered, encoding="utf-8")
        print(f"wrote {target} ({len(calls)} calls)")
    if stale:
        print(
            "regenerate with: python3 tools/mcp-server/smoke/"
            f"{Path(__file__).name}", file=sys.stderr)
    return 1 if stale else 0


if __name__ == "__main__":
    raise SystemExit(main())
