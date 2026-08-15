# Epic #435 acceptance smoke — agent-safe headless document editing

Run 2026-08-15 against `main` at `400af83` plus the three commits this run produced.
Transport: `docxodus-mcp` over stdio JSON-RPC, driven by
[`tools/mcp-server/smoke/mcp_probe.py`](../../tools/mcp-server/smoke/mcp_probe.py).

Reproduce with the commands in [`tools/mcp-server/smoke/README.md`](../../tools/mcp-server/smoke/README.md).

## Document

`TestFiles/NVCA-Model-COI.docx` — a model certificate of incorporation, chosen as the
in-repo equivalent of the certificate the round-three smoke used.

| Property | Value |
|---|---|
| OPC parts | 44 |
| Running stories | 8 headers, 10 footers |
| Notes | 94 footnotes |
| Bookmarks | 392 (body) |
| Anchors | 462 |
| Blocks | 153 `h`, 59 `p`, 22 `li`, 4 `sec` |
| Section | US Letter, `lowerRoman` page numbers, `first` header + `even`/`default`/`first` footers |
| Tables / images / hyperlinks / content controls / revisions / comments | none at open |

## Result

| Pass | Calls | Assertions | Failed | Expected failures | Replay mismatches |
|---|---:|---:|---:|---:|---:|
| Workflow | 45 | 68 | **0** | 4 | **0** |
| Reopen validation | 13 | 27 | **0** | 0 | — |

Unit suite after the fix below: **3793 passed, 0 failed, 3 skipped**.

## The spine: inspect → preview → apply → retry → audit

**Inspect.** Anchors were discovered by content, never assumed: text search for the
operative name clause and the certification clause, kind search for a list item. Section
introspection returned the full `sectPr` rollup including per-kind header/footer refs.
`docxodus_pagination status` correctly reported `unavailable` / `no_page_map` before any
map was registered.

**Preview.** A six-step tracked batch (surgical placeholder fill, paragraph insert,
footnote, comment, format) previewed under `mode: preview` with a batch-start
`expectedVersion` guard and `previewHtml: full`.

| Predicted | Value |
|---|---|
| `baseVersion` → `resultVersion` | 0 → 1 |
| `revisionChanges.added` | 4 |
| `commentChanges.added` | 1 |
| `packageHash` | `ec364b6d…` |
| rendered HTML | 522,302 bytes |

The live session was then re-inspected: version still 0, anchors still 462, footnotes still
94, zero revisions, zero comments. The isolated dry run of #446 holds — nothing crossed
from the shadow.

**Apply.** The same six steps under `mode: atomic` with
`transactionId: epic435-smoke-restated-certificate`. `resultVersion` was asserted *against
the preview's captured prediction*, not a literal, and matched at 1; `revisionChanges.added`
and `commentChanges.added` matched the predictions exactly.

`packageHash` differed between preview and apply (`ec364b6d…` vs `64ce4960…`), which is the
documented contract, not a defect: this batch generates a comment id, a footnote id and
revision timestamps, and both receipts carry the warnings that say so —

> Created anchors and related OOXML ids may be generated independently on replay;
> preview/apply equivalence is semantic and packageHash or anchor ids may differ.

Equivalence was therefore verified where it is promised — outcomes, versions, and semantic
deltas — and not asserted where the receipt says not to.

**Retry.** The identical request was re-sent to simulate a lost response. The runner
compared the raw wire text of the two responses: **byte-exact**, `replayMismatches: 0`.
Nothing re-applied — version, anchor count, footnote count, revision count and comment
count were all unchanged after the retry.

**Audit.** Version, anchors and revisions were re-checked after every guarded step; undo
and redo both remained usable and returned the document to the same anchor count.

## Failing closed

Four calls are expected to fail, and the harness treats a success there as the defect.

| Probe | Result |
|---|---|
| `add_bookmark` while tracking | `tracked_operation_unsupported` — bookmark mutations have no faithful tracked encoding |
| `insert_table` while tracking | `tracked_operation_unsupported` — no reversible tracked encoding on this document shape |
| Stale `expectedVersion` | `precondition_failed`, reporting the true `currentVersion` |
| Atomic batch with a bad anchor at index 2 | `status: failed`, `rolledBack: true`, `editsApplied: 0`, `failure.index: 2`, `failure.error.code: anchor_not_found` |

After the rolled-back batch: version, anchor count and revision count all unchanged, and
neither rolled-back string appears anywhere in the package. A bookmark endpoint aimed
inside the tracked insertion was separately refused with `unsupported_inline_boundary`.

## Capability coverage

Exercised: paragraphs, headings, lists, footnotes, bookmarks, hyperlinks, tables,
comments, tracked revisions, run and paragraph formatting, section introspection, page
citations, undo/redo, atomic batches, preconditions, transaction identity.

Link and table work runs in its own untracked batch, reached with
`docxodus_track_changes set_mode` — the engine refuses those families under tracked
recording rather than emitting markup it cannot reverse, and switching mode mid-workflow is
the supported path (#304).

Not applicable, with rationale:

| Capability | Rationale |
|---|---|
| Images | The document contains none, and there is no image the workflow would legitimately author into a charter. The surface is covered by `MCP145`. |
| Content controls | The document contains none, and the `#452` surface fills existing controls rather than creating them. Covered by `MCP147`/`MCP148`. |
| Endnotes | The document uses footnotes only. |

Page citations were exercised by registering a page map and citing an anchor against it.
The map is caller-supplied by design — the renderer is client-side — so a synthetic
single-page map was registered and the citation round-tripped with the anchor's fragment
geometry and `pageName: "i"` (matching the document's `lowerRoman` numbering).

## Persistence

Saved with `persistAnchorIds: true`, closed, reopened, and re-inspected.

- **Package integrity**: zip valid, 44 → 45 parts, **no part lost**, the only addition
  `word/comments.xml`. No dangling relationships; exactly one external relationship (the
  authored hyperlink).
- **Persisted content**: 4 revisions in document order (insert paragraph, delete
  placeholder, insert name, format name), all authored `Smoke Counsel`; 1 comment; the
  bookmark resolves to its anchor; the hyperlink persists as an external relationship with
  its target intact; the table reopens with an addressable 2×2 grid; the footnote text is
  findable.
- **Absence of unintended change**: neither rolled-back string nor the refused
  stale-precondition string appears anywhere in the reopened package.

## Defects found

### 1. Tracked insertions were invisible to every offset-addressed surface — FIXED

Under `render_inline`, an agent could not re-find the edit it had just made.
`replace_text_range` succeeded and the projection showed the new text, but
`docxodus_search` returned nothing, `apply_format_by_substring` reported
`offset_out_of_range`, and a follow-up `replace_text_range` could not address it.

Root cause: `InlineRuns` descends only into `InlineContainerNames`, which omitted `w:ins`,
so every run inside a tracked insertion was missing from the flat string all
offset-addressed ops share — while `Project().Markdown` and `TextPreview` included it.

Not limited to text the session authored: `RA001-Tracked-Revisions-01.docx` holds six
occurrences of "chromatogram" and search returned exactly the two outside `w:ins`, so an
incoming Word redline had its inserted spans silently skipped too.

Fixed in `e4b4365` by making `w:ins`/`w:moveTo` transparent containers and leaving
`w:del`/`w:moveFrom` opaque — the split is visible text versus text the document says was
removed, not plain versus revised runs. Coverage: `DS409`, `DS410`.

### 2. A batch step that matches nothing fails as `internal_error` — OPEN

`replace_text_range` with a substring that is genuinely absent behaves inconsistently:

- called directly, it returns `[]` — no error, no signal, indistinguishable from a
  successful no-op;
- inside a `docxodus_mutations` batch, that same empty result becomes
  `internal_error: "batch mutation returned no valid edit results"`, failing and rolling
  back the whole batch.

"No match" is an ordinary, expected outcome and deserves a structured, actionable code on
both paths. Deciding whether the direct call should start reporting it is a public-API
change, so this is filed rather than fixed here.

## Gate status

Every success criterion in the epic's final acceptance gate is met by the two runs above,
with defect 1 resolved and defect 2 recorded. Defect 2 does not block the gate: it is an
error-classification inconsistency on an operation that matches nothing, and no step of the
acceptance workflow depends on it.
