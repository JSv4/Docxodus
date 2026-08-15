# DOCX Mutation API

> **Status:** Implemented. Source: `Docxodus/DocxSession.cs`, `Docxodus/RawDocxOps.cs`, `Docxodus/Internal/MarkdownPayloadParser.cs`, `Docxodus/Internal/UndoRing.cs`. Tests: `Docxodus.Tests/DocxSessionTests.cs` (`DS###`), `MarkdownPayloadParserTests.cs`, and an end-to-end smoke at `DocxSessionSmokeTest.cs`. WASM bridge: `wasm/DocxodusWasm/DocxSessionBridge.cs`. npm wrapper: `npm/src/session.ts`. The full type-level spec lives at `docs/superpowers/specs/2026-05-24-docx-mutation-api-design.md` — this doc is the conceptual reading and the recipe book; it points to source for the canonical shapes rather than restating them.

## What this is

`DocxSession` is the **write-side counterpart** to `WmlToMarkdownConverter`. The projector turns a DOCX into anchor-addressed markdown; the session lets you mutate the same DOCX by those anchor ids — replace text, insert/split/merge paragraphs, apply formatting, edit table cells — without the agent (or human) ever having to think about OOXML. Anything the markdown subset can't express drops to a clearly-namespaced raw-XML escape hatch.

The intended consumer is an agentic editing pipeline: an LLM reads the markdown projection of a document, decides what to change, and calls a small set of high-level tools. But the same surface is useful for any tooling that wants to make surgical, ID-addressed edits to Word documents — review pipelines, structured-edit UIs, templating workflows.

## Why it's shaped the way it is

Three design forces, in order of weight:

**The agent must not learn OOXML.** Every public method takes an anchor id (a string) and either a markdown payload (a string) or a small typed value (a `FormatOp`, a `CharSpan`). The agent never sees an `XElement`, never picks an SDK type, never has to know that bold is `w:b` inside `w:rPr`. The Raw escape hatch exists for the cases the markdown subset can't reach, but it's a separate namespace (`session.Raw.*`) so it's syntactically obvious when you've left the safe zone.

**Edits must be reversible.** Agents make mistakes. The session keeps a bounded ring of pre-op snapshots (default 20 deep) so `Undo()` and `Redo()` work without the caller orchestrating anything. Ordinary single-op snapshots are per-part XML clones; an explicit transaction uses a complete package checkpoint because a batch can change arbitrary parts and relationships.

**Errors must be pattern-matchable, not stringly-typed.** Every mutation returns an `EditResult` envelope; failure carries a typed `EditErrorCode` with a remediation message. The same enum is exposed as a snake-case string union in TypeScript, so JS agents pattern-match the same way C# callers do. No method on the session throws across the boundary (the constructor and `Save()` are the only places that can — and only for fatal conditions like an invalid DOCX or IO failure).

## Document version and optimistic preconditions

Every session exposes a monotonic `long Version`. It is `0` when the document is
opened and advances exactly once for each committed document mutation. A
multi-match `ReplaceTextRange` is one mutation. `Undo()` and `Redo()` also advance
the version—restoring older content never restores an older caller-visible
version. Validation failures, precondition failures, exceptions that roll back,
and successful no-ops do not advance it. Version state is carried in internal
snapshots so rollback and speculative preview restoration cannot leak a version
change.

Callers that derived an edit from an earlier projection can guard the mutation:

```csharp
var info = session.GetAnchorInfo(anchor)!;
var result = session.ExecuteMutation(
    new MutationPreconditions
    {
        ExpectedVersion = session.Version,
        AnchorId = anchor,
        ExpectedContentHash = info.ContentHash,
        ExpectedText = info.VisibleText,
        ExpectedKind = info.Kind,
        ExpectedScope = info.Scope,
    },
    s => s.ReplaceText(anchor, replacement));
```

`MutationPreconditions` fields are optional and ANDed:

| .NET | wire | Meaning |
|---|---|---|
| `ExpectedVersion` | `expectedVersion` | The session version used to derive the plan. |
| `AnchorId` | `anchorId` | Guard target. Mutation façades infer their primary target when omitted. |
| `ExpectedContentHash` | `expectedContentHash` | Hash of the target's current OOXML subtree, also returned by `GetAnchorInfo`. |
| `ExpectedText` | `expectedText` | Exact current visible text, also returned as `AnchorInfo.VisibleText`. |
| `ExpectedTextRange` | `expectedTextRange: {start,length,text}` | Exact ordinal substring guard over visible text. |
| `ExpectedKind` / `ExpectedScope` | `expectedKind` / `expectedScope` | Current canonical anchor metadata. A stale kind prefix still resolves by Unid and reports the new kind. |
| `ExpectedMatchCount` | `expectedMatchCount` | Exact live occurrence count for `ReplaceTextRange`. |

`EvaluatePreconditions`/transport `checkPreconditions` are read-only probes.
`ExecuteMutation` is the direct .NET gated primitive used by the shared façade.
`ReplaceTextRange` holds the same gate across initial guard evaluation, live match
enumeration/counting, and every replacement, so another mutation cannot slip
between count and commit.

On mismatch, no document bytes or history entry change and the result has code
`PreconditionFailed` (`precondition_failed` on the wire). Its `precondition` member
contains `condition`, `expected`, `actual`, `currentVersion`, and `currentTarget`
(`exists`, canonical anchor id/kind/scope, content hash, and exact visible text).
That is enough for an agent to decide whether to rebase, retarget, or abandon the
edit without an extra diagnostic round trip.

The common wire object is available throughout the stack: npm exposes
`getVersion`, `checkPreconditions`, and `runWithPreconditions`; Python exposes
`get_version`, `check_preconditions`, and a `session.preconditioned(...)` context
that attaches the guard to each mutation request; stdio accepts top-level
`preconditions`; MCP mutation tools and individual batch steps accept the same
property. MCP batches may additionally carry a batch-start guard. Preview mode
restores the starting version after it undoes its speculative edits.

## Atomic batches and transactions

`ExecuteBatch(steps, mode)` is atomic by default. Each `MutationBatchStep` names a
tool/action for diagnostics, supplies a synchronous mutation callback, and may
provide a read-only preflight callback:

```csharp
var result = session.ExecuteBatch(new[]
{
    new MutationBatchStep("docx_edit", "replace_text",
        s => s.ReplaceText(firstAnchor, "First replacement")),
    new MutationBatchStep("docx_create", "set_header_text",
        s => s.SetHeaderText(firstAnchor, HeaderFooterKind.Default, "Confidential")),
});
```

Atomic mode evaluates every available preflight against the batch-start state
before step zero. A successful batch is one caller-visible version advancement and
one undo/redo unit, regardless of its step count. A failed result or thrown step
restores document content, all part/relationship topology, annotations/custom XML,
anchor and revision generators, mutable tracking configuration, version, and both
history cursors. The structured failure identifies `index`, `tool`, `action`,
`error`, and `rolledBack`; failure consumes no history and advances no version.

`MutationBatchMode.BestEffort` must be selected explicitly. It preserves partial
successes and evaluates each step's preflight immediately before that step, so a
later preflight can observe state made by an earlier successful mutation.

`BeginTransaction()` exposes the same full-package checkpoint for façade code that
needs callback composition. Transactions are synchronous, same-thread, and strict
LIFO, but nested scopes are supported: inner commits remain speculative, and the
outer commit squashes all nested work into one history/version unit. Dispose without
`Commit()` rolls back. A validation failure from wrong-thread or out-of-order
completion leaves the scope active and recoverable. Disposing the owning session on
the owner thread abandons all scopes and releases their mutation-gate entries.

The same semantics reach `DocxSessionOps`/JSON, WASM and npm
(`session.executeBatch`), stdio and Python (`session.execute_batch`), and MCP
(`docxodus_mutations`).

### Isolated previews

`PreviewBatch(steps, mode, options)` runs the identical step delegates against a complete
clone of the live package (`CreateShadowSession`). Guards, mutations, history writes,
semantic inspection, package hashing, and any HTML render all target the shadow, which is
disposed on every return and throw path. Abandoning a preview therefore cannot require a
live rollback — the live session's bytes, caches, version, configuration and undo/redo
cursors were never mutation targets.

```csharp
var preview = session.PreviewBatch(
    new[]
    {
        new MutationBatchStep("docx_edit", "replace_text",
            s => s.ReplaceText(firstAnchor, "Proposed replacement")),
    },
    options: new MutationBatchPreviewOptions
    {
        HtmlMode = MutationPreviewHtmlMode.Full,
    });
```

**Caller contract for the typed overload.** The `s` argument each callback receives IS the
shadow. Isolation comes from addressing that argument; a callback that closes over the live
session and calls `liveSession.ReplaceText(...)` instead mutates the live document, and
nothing in the typed API can prevent it. The handle-shaped seams — `DocxSessionOps.PreviewBatch`
and the `OpenPreviewSession` bridge — are intrinsically safe because a step factory there is
handed only the temporary shadow handle and never sees the live one.

`MutationBatchPreviewOptions.HtmlMode` (`MutationPreviewHtmlMode`: `None`, `Scoped`, `Full`;
`Scoped` additionally requires `HtmlAnchorId`) renders the predicted document from the shadow.
The option profile lives in exactly one place — `HtmlConversionOps.PreviewDocumentOptions()`
and `PreviewBlockOptions()` — and every surface consumes it, so a browser preview and an
MCP preview of the same batch describe the same document. A preview shows tracked changes,
comments, annotations, notes and headers/footers: it answers "what would this document
become", which is not the editor's authoring view (`DocxSessionOps.RenderHtml`, where
comments and annotations are off).

Both preview and apply return the same enriched receipt: `baseVersion`/`resultVersion`,
`packageHash`, `{added, removed, modified}` change sets for revisions, comments and
annotations, and `warnings`. Change-set membership is decided on each entry's SERIALIZED
projection — the shape the transports actually publish, and the same comparison npm makes —
never on CLR equality. `packageHash` is `null`, never `""`, when it could not be computed;
an absent hash must not compare equal to another absent hash.

**Cost.** Enrichment is unconditional on BOTH paths. Every batch — applied or previewed —
runs `ListRevisions` + `ListComments` + `ListAnnotations` twice (before and after, each
forcing an anchor index) plus one `GetPackageContentHash()`, which serializes a full package
checkpoint and SHA-256s it. A preview additionally pays a package clone and a second open
`WordprocessingDocument`, roughly doubling peak memory for the duration. That is material on
a constrained heap (a browser WASM session holding a large DOCX). There is deliberately no
opt-out today; gating it behind a setting is a public-API decision that has not been taken.

npm exposes this as `session.previewBatch(steps, mode, { html, htmlAnchorId })` with callbacks
that receive the shadow session, stdio/Python as `session.preview_batch(steps, mode,
html_mode=…, html_anchor_id=…)`, and MCP as `docxodus_mutations` with `"mode": "preview"`.

Two properties of the wire-serialized transports (MCP and stdio, which round-trip
each step's `EditResult` through `DocxSessionJson`) are worth stating explicitly:

- **Step receipts are lossless.** A batched structural table op keeps its
  `tableAnchors` mapping, so a caller can address the rows/columns/cells the same
  batch just created. `DocxSessionJson.ParseEditResult` is the inverse of
  `Serialize(EditResult)`; adding a field to one without the other silently deletes
  it from every batch receipt.
- **`expectedMatchCount` is evaluated late.** Every other guard is decided by the
  read-only preflight — at the batch-start boundary in atomic mode. A match count
  can only be evaluated by an op that has enumerated the live matches, and
  `ReplaceTextRange` is the sole supplier, so that one guard is carried into the
  step and enforced at the step's own turn. A batch mixing an `expectedText` and an
  `expectedMatchCount` guard therefore compares them against two different states.

## Architecture

```
┌──────────────────────────────────────────────────────────────┐
│  npm: openDocxSession(bytes, settings?) → DocxSession        │
│       session.replaceText(anchor, md) → EditResult           │
│       session.undo(); session.save() → Uint8Array            │
└──────────────────────────────────────────────────────────────┘
                            │  JS ↔ WASM (handle = int sessionId)
┌──────────────────────────────────────────────────────────────┐
│  wasm/DocxodusWasm/DocxSessionBridge.cs                      │
│    static Dictionary<int, DocxSession> _sessions             │
│    [JSExport] static methods, JSON-serialized in/out         │
└──────────────────────────────────────────────────────────────┘
                            │
┌──────────────────────────────────────────────────────────────┐
│  Docxodus/DocxSession.cs            (the real work)          │
│    sealed class DocxSession : IDisposable                    │
│      - long-lived WordprocessingDocument over a MemoryStream │
│      - tier A/B/C/D mutation methods + Raw escape hatch      │
│      - UndoRing<DocumentSnapshot> for bounded undo           │
└──────────────────────────────────────────────────────────────┘
                            │ owns
        ┌───────────────────┼────────────────────────┐
        ▼                   ▼                        ▼
  WordprocessingDoc   AnchorIndex            UndoRing
  (live XDocument     (refreshed lazily      (per-part XML
   per part)           after each mutation)   snapshots, default 50)
```

The session owns one `WordprocessingDocument` open over its own `MemoryStream`. Mutations operate directly on the in-memory `XDocument` of the affected part. Re-projection uses the existing `WmlToMarkdownConverter` over the live document.

For the full public surface — exact method signatures, settings, value types — read `Docxodus/DocxSession.cs` end-to-end. It's ~700 lines and organized by tier.

## How to think about anchors

An anchor id looks like `{#h:body:7b9f61007f9341c8aa5878ee63ffc874}`. The parts:

- `kind` — what kind of OOXML element this is (`p`, `h`, `li`, `tbl`, `tr`, `tc`, `cmt`, `fn`, `en`, `img`, `drw`, `unk`).
- `scope` — which package part it lives in (`body`, `hdr1`/`hdr2`/…, `ftr1`/…, `fn`, `en`, `cmt`).
- `unid` — a 32-char hex stable identifier (Docxodus's `PtOpenXml.Unid`).

**The Unid is the identity.** The `kind:scope:` prefix is descriptive metadata and can change across mutations. Promoting a `Normal` paragraph to `Heading2` flips its anchor id from `{#p:body:abcd}` to `{#h:body:abcd}`. The session's lookup helper (`DocxSession.FindAnchor`) does a direct dictionary hit first, then falls back to a Unid-only scan, so a cached id whose prefix has gone stale still resolves. Even so, prefer the `Modified` entry returned in each `EditResult` for the current canonical form — the fallback is cheap insurance, not a long-term substitute for tracking renames.

**Created/Removed/Modified are the contract.** Each mutation returns three anchor lists in its `EditResult`. The lifecycle policy is documented in the [Anchor lifecycle](#anchor-lifecycle) section below — that's the contract the agent's mental model is supposed to track.

## What the markdown payload subset is, and why

When you pass markdown into `ReplaceText`, `InsertParagraph`, or `ReplaceCellContent`, the session runs it through `MarkdownPayloadParser`, a hand-rolled parser that accepts **only** what the projector emits. Block-level: paragraphs, ATX headings (`#`–`######`), bulleted lists, ordered lists (with indent-based nesting), blockquotes, fenced code blocks. Inline: `**bold**`, `*italic*`, `` `code` ``, `~~strike~~`, `[text](url)` links, the GFM hard break (`"  \n"`, two trailing spaces → a real `w:br`), backslash escapes. (A *blank* line still separates paragraphs; a single in-paragraph newline becomes a `w:br`, symmetric with the projector's `w:br → "  \n"`.)

This is symmetric by design: anything the projector can emit, the parser can accept, so an agent can read markdown out and write markdown in. Anything outside the subset is rejected with a typed error that names either the v1 op to use instead or the v2 op planned to address it. The full table of accepted and rejected syntax is in the spec — the practical shorthand:

- If you can see it in the projection output, you can write it in a payload.
- If you need a table → `InsertTable(anchor, Position, rows, cols, TableInsertOptions?)` (borderless, row-major `CellContents`, `CellAlignment`, per-column `ColumnWidths`). It returns canonical `tc` anchors. Discover the whole shape with `GetTableMetadata(tblAnchor)` or translate in either direction with `ResolveTableCellAnchor(tcAnchor)` / `ResolveTableCellCoordinate(tblAnchor, row, column)`. Every cell-content, shape, merge, and styling operation takes that same canonical `tc` anchor: `ReplaceCellContent`, `InsertTableRow`/`InsertTableColumn`/`DeleteTableRow`/`DeleteTableColumn`, `SetColumnWidths`, `SetTableBorders`, `SetCellShading`, `SetRepeatHeaderRow`, `SetTableRowOptions`, `MergeCells`, and `UnmergeCells`. See [Canonical table addressing](#canonical-table-addressing) and [the grid model](#table-cell-merge-the-grid-model).
- If you need a footnote or endnote → `InsertFootnote(anchor, offset, markdown)` / `InsertEndnote(...)`; a `[^label]` reference in a *payload* stays rejected, because a label can't name a note the payload doesn't define.
- If you need a comment → `AddComment(anchor, span?, author, markdown, initials?, date?)`, or target a tracked change from `ListRevisions()` with `AddCommentToRevision(revisionId, author, markdown, initials?, date?)`; reply with `AddCommentReply(parentCmtAnchor, author, markdown, initials?, date?)`, and resolve/reopen with `SetCommentResolved(cmtAnchor, resolved)`. A `{#cmt:...}` token in a *payload* stays rejected, because inline comment tokens are projection output only (see the Comments section).
- If you need an image → `InsertImage(anchor, offset, imageBytes, ImageInsertOptions?)`, with `ListImages`/`ReplaceImage`/`SetImageDimensions`/`SetImageMetadata`/`SetImageFloatingLayout`/`RemoveImage` alongside it (see [native_images.md](native_images.md)). An `![alt](url)` in a *payload* stays rejected, because a picture needs binary bytes and the parser has no way to fetch a URL.
- For everything OOXML can do that markdown can't (complex tables, math, content controls, drawings) → `session.Raw.*`.

We didn't pick CommonMark or GFM as the input language because the projector's subset is small and well-defined; running a full parser against that subset would import surprise (e.g., GFM tables silently splitting paragraphs, autolinks mis-classifying spans). The hand-rolled parser is ~300 LOC, has no dependencies, and gives us complete control over what gets rejected and why.

Two round-trip quirks worth knowing when you write tests against the markdown output:

- **The projector escapes markdown punctuation in text content.** `-`, `*`, `_`, `` ` ``, `~`, `\`, and other characters that could be parsed as markdown are backslash-escaped (e.g., `RAWSIBLING-INSERTED` projects as `RAWSIBLING\-INSERTED`). Don't write literal `Contains(...)` assertions over hyphenated tokens; either strip backslashes from the projection or use tokens without markdown-significant characters.
- **`InsertParagraph` with a bulleted markdown payload does not inherit list numbering.** A payload like `- item one` parses as a `BulletItem` block and the created anchor has kind `li`, but the inserted paragraph has no `w:numPr`, so Word renders it as a plain paragraph (and `SetListLevel` will return `AnchorWrongKind` because there is no numbering to adjust). To get a real bulleted item in v1, use `Raw.InsertXml` with a fragment derived from `Raw.GetXml(existingListItemAnchor)` so the `w:numPr` and numbering id come along for free. A first-class numbering-inheritance path is on the v2 list (see Known limits).

## Anchor lifecycle

Each mutation reports which anchors it created, removed, or modified. This table is the contract — agent harnesses use it to keep their cached projection in sync without re-projecting on every call:

| Op | Created | Removed | Modified | Patch scope |
|---|---|---|---|---|
| `ReplaceText(p, md)` | — (markdown subset can't introduce inline anchors in v1) | descendant inline anchors that no longer exist (rare) | `p` | `p` |
| `DeleteBlock(p)` (or `h`/`li`/`tbl`) | — | `p` + all descendant anchors | — | nearest stable ancestor |
| `DeleteBlock(fn)` / `DeleteBlock(en)` / `DeleteBlock(cmt)` | — | the definition anchor (and any cross-references it pointed at — those become "gone" but aren't separately addressed) | — | nearest stable ancestor in the body |
| `InsertParagraph(p, pos, md)` | one anchor per new block | — | — | smallest enclosing common parent |
| `SplitParagraph(p, offset)` | the **second** half | — | `p` (first half — convention) | enclosing parent |
| `MergeParagraphs(a, b)` | — | `b` + descendants | `a` | `a` |
| `ApplyFormat(p, span?, op)` | — | — | `p` | `p` |
| `SetParagraphStyle(p, style)` | — | — | `p` (kind prefix may flip) | `p` |
| `SetListLevel(p, delta)` | — | — | `p` | enclosing list (downstream items renumber) |
| `RemoveListMembership(p)` | — | — | `p` (kind flips `li`→`p`) | enclosing list |
| `ApplyListFormat(p, fmt)` | — | — | `p` (kind may flip `p`↔`li`) | `p` |
| `ApplyListFormatRange(first, last, fmt)` | — | — | every `w:p` member of the run (kinds may flip `p`↔`li`) | `first` |
| `SetListStartOverride(li, value)` | — | — | the anchored item + every following member of its numbering instance (all repointed to a dedicated `w:num`) | the anchored item |
| `ClearListStartOverride(li)` | — | — | every member of the item's numbering instance (all repointed together) | the anchored item |
| `ReplaceCellContent(tc, md)` | — | descendant inline anchors (rare) | `tc` | `tc` |
| table row/column CRUD, merge/unmerge | new `tr`/`tc` identities where applicable | invalidated `tr`/`tc` identities | addressed `tc` where applicable | enclosing `tbl`; full structural map in `TableAnchors` |
| `SetHeaderText(p, kind, md)` / `SetFooterText(...)` | the new header/footer paragraph anchors (scope `hdr{N}`/`ftr{N}`) | — (reused-part old paragraphs cease to exist; not separately reported in v1) | — | whole document |
| `InsertPageNumberField(p, field?)` | — | — | `p` (the paragraph the field is appended to) | `p` |
| `InsertFootnote(p, offset, md)` / `InsertEndnote(...)` | the note definition (`fn`/`en`) + its paragraphs (scope `fn`/`en`) | — | `p` (the citing paragraph) | whole document |
| `AddComment(p, span?, author, md, …)` | the comment definition (`cmt`) + its paragraphs (kind `p`, scope `cmt`) | — | `p` (the commented paragraph) | `p` |
| `AddCommentToRevision(revisionId, author, md, …)` | the comment definition (`cmt`) + its paragraphs (kind `p`, scope `cmt`) | — | every block touched by the revision | whole-document re-projection |
| `AddCommentReply(cmt, author, md, …)` | the reply definition (`cmt`) + its paragraphs (kind `p`, scope `cmt`) | — | the parent `cmt` plus every document-side `p` hosting its reference | the first referenced host `p` (normally the sole host) |
| `UpdateComment(cmt, md)` | the new body paragraph anchors (scope `cmt`) | the old body paragraph anchors | `cmt` | `cmt` |
| `SetCommentResolved(cmt, resolved)` | — | — | `cmt` | `cmt` |
| `RemoveComment(cmt)` | — | `cmt` + descendant paragraph anchors (the `DeleteBlock(cmt)` shape) | — | nearest stable ancestor |
| `Raw.InsertXml(a, pos, xml)` | every block in the new XML | — | — | enclosing parent |
| `Raw.ReplaceXml(a, xml)` | unids present in the new XML but not the old (typical for caller-authored XML) | unids present in the old element but not the new (when `a` itself is gone) | unids present in both (typical for the `GetXml → mutate → ReplaceXml` round trip, which preserves Unids) | enclosing parent |
| `Undo()` / `Redo()` | (diff vs current) | (diff vs current) | (diff vs current) | `null` — caller re-projects |

Two conventions worth pinning down because they affect agent reasoning:

- **`SplitParagraph` keeps the original Unid on the first half.** Reason: external systems (LLM context windows, search indices) bias toward the pre-split anchor position; keeping the prefix-half stable minimizes invalidation downstream.
- **`MergeParagraphs` lets the first anchor absorb the second.** Symmetric reason: the first anchor is to the left in reading order and is more likely to be the one a caller has cached.

**Tracked-change mode shifts the semantics for `ReplaceText` and block deletion (`DeleteBlock`, `DeleteRange`, and `DeleteSection`).** When `Settings.TrackedChanges = RenderInline`, supported deletions don't remove elements — they wrap old runs in `w:del` and new content in `w:ins`. So the affected anchor stays live and appears in `Modified` instead of `Removed`. The agent's view of the world doesn't have to change; the `EditResult` shape is unchanged. The mode is switchable mid-session — see "Switching tracked-changes mode mid-session" below.

Structural tracking is deliberately capability-gated. Row insertion/deletion and column
insertion emit native Word row/cell/property revisions; single-paragraph list application,
removal, and level changes emit `numPr/w:ins` or `pPrChange`. Shapes without a safely reversible
encoding—tracked column deletion, merge/unmerge, range list formatting, and list-start
overrides—return `TrackedOperationUnsupported` without mutation or history. A table with a live
cell structural revision returns `UnresolvedStructuralRevision` before another structural edit.

**`ReplaceText` quietly strips a leading auto-number prefix from the payload.** When the target paragraph carries `w:numPr` (numbered heading or list item), the projector emits the resolved number inline (`## Fourth The total number…`) so a human can read what Word renders. An agent that echoes the visible heading back as its `ReplaceText` payload would otherwise see `Fourth Fourth: …` in the saved DOCX — the auto-number is still applied by Word, *and* the new run text now also starts with the prefix. The session resolves the number via the shared `Internal.ListNumberResolver` and strips a matching prefix (plus one optional separator: space, tab, or NBSP) from the payload before parsing. Idempotent — if the agent skipped the prefix, nothing is stripped. Documented in `DS091`/`DS091b`.

## When to use what

Decision tree for the agent (or its prompt):

```
What am I editing?
├── Just the visible text of a paragraph/heading/list item?
│       → ReplaceText(anchor, markdown)
│
├── Removing a paragraph/heading/list item?
│       → DeleteBlock(anchor)
│
├── Adding a paragraph adjacent to an existing one?
│       → InsertParagraph(anchor, "before" | "after", markdown)
│
├── Splitting one paragraph into two?
│       → SplitParagraph(anchor, offset)   # offset is character position
│
├── Joining two adjacent paragraphs?
│       → MergeParagraphs(firstAnchor, secondAnchor)
│
├── Just the bold/italic/underline/code/color/size/font of some characters?
│       → ApplyFormat(anchor, CharSpan(start, length), FormatOp{...})
│       → ApplyFormat(anchor, null, FormatOp{...})  # null span = whole paragraph
│         # FormatOp.FontSizePts → w:sz/w:szCs; FontFamily → w:rFonts ("" clears)
│
├── Changing a paragraph's style (e.g., Normal → Heading2)?
│       → SetParagraphStyle(anchor, styleId)
│
├── Paragraph layout — alignment, indents, spacing, page-break-before, borders?
│       → SetParagraphFormat(anchor, ParagraphFormatOp{...})
│         # All indent/spacing values are twips (1440 = 1in, 20 = 1pt). IndentDelta shifts
│         # w:ind/@w:left relatively; FirstLineIndent/HangingIndent are absolute and share
│         # one either/or w:ind slot (setting one evicts the other; both in one op →
│         # InvalidParagraphFormat). SpacingBefore/SpacingAfter → w:spacing/@w:before/@w:after.
│         # LineSpacing → w:spacing/@w:line, measured per LineSpacingRule: auto (default) =
│         # 240ths of a line (240 single, 360 = 1.5×, 480 double), exact/atLeast = twips.
│
├── Indenting/outdenting a list item or removing it from a list?
│       → SetListLevel(anchor, +1 | -1)
│       → RemoveListMembership(anchor)
│
├── Making a paragraph a real auto-numbered list item (or changing a list's format)?
│       → ApplyListFormat(anchor, ListFormat.Bullet | Decimal | LowerLetter | UpperLetter |
│                                  LowerRoman | UpperRoman | *Parenthesis variants | None)
│         # Synthesizes a reusable numbering definition if needed (NumberingFactory, marker
│         # nsid per format). Plain formats render "1."/"a."/"i." level text; the *Parenthesis
│         # variants render "(1)"/"(a)"/"(i)" — same w:numFmt, different w:lvlText (the
│         # legal-drafting presets). None strips inline list membership.
│       → ApplyListFormatRange(firstAnchor, lastAnchor, format)
│         # The same conversion across a contiguous sibling run, first..last INCLUSIVE (either
│         # document order). One call instead of one per item; every member is guaranteed the
│         # same shared w:num instance so the sequence numbers stay intact; each member keeps
│         # its own w:ilvl; the whole range is ONE undo step. Non-paragraph siblings inside
│         # the range are skipped. Anchors in different parts or different direct parents →
│         # AnchorsNotAdjacent.
│
├── Restarting (or seeding) a list's numbering — Word's "Set Numbering Value…"?
│       → SetListStartOverride(itemAnchor, value)   # "start this list at 5"
│         # Writes w:lvlOverride/w:startOverride on a DEDICATED w:num (the item's current
│         # instance is cloned, never mutated — it may be shared, and the numbering part is
│         # not snapshotted) and repoints the anchored item + every FOLLOWING member of its
│         # instance. Anchoring mid-list therefore splits the sequence exactly like Word:
│         # earlier items keep their numbers, the tail continues from `value`. Negative
│         # value → InvalidListStartValue. Requires a list-item ("li") anchor.
│       → ClearListStartOverride(itemAnchor)        # back to "continue from the definition"
│         # Repoints EVERY member of the instance (they move together, so relative
│         # continuation is preserved) at a clone WITHOUT the override; a sequence with no
│         # override at the item's level is a successful, undo-free no-op.
│
├── Replacing the contents of a table cell?
│       → ReplaceCellContent(tcAnchor, markdown)
│
├── Setting a section's running header/footer (any body block anchors the section)?
│       → SetHeaderText(bodyAnchor, HeaderFooterKind.Default|First|Even, markdown)
│       → SetFooterText(bodyAnchor, ...)   # Created lists the hdr{N}/ftr{N} paragraphs
│
├── Putting a page number in a header/footer paragraph?
│       → InsertPageNumberField(hdrOrFtrAnchor, PageNumberField.CurrentPage|TotalPages)
│
├── Adding a footnote/endnote (cited from a body paragraph at a character offset)?
│       → InsertFootnote(bodyAnchor, offset, markdown)
│       → InsertEndnote(bodyAnchor, offset, markdown)   # Created lists the fn/en anchors
│   …then edit it with ReplaceText(notePara, md) or drop it with DeleteBlock(noteDef)
│
├── Inserting/deleting table rows or columns, merging cells,
│   embedding a chart, inserting a math equation,
│   adding a content control?
│       → Drop to session.Raw.*  (v2 ops planned for the common cases)
│
└── Anything that needs an undo guard?
        → Just call it. Every successful op takes a snapshot.
          session.Undo() restores prior state.
```

## Block reordering — `MoveBlock` and `ValidMoveTargets`

```csharp
EditResult MoveBlock(string sourceAnchorId, string targetAnchorId, Position pos);
IReadOnlyList<MoveTarget> ValidMoveTargets(string sourceAnchorId);  // record(AnchorId, Before, After)
```

`MoveBlock` reorders ONE top-level body block relative to another — the model half of the
editor's drag handle. Source and target must be block kinds (`p`, `h`, `li`, `tbl`), live in
the same part, and share a direct XML parent (so a block-level `w:sdt` boundary is never
crossed: that would change content-control membership). Moving a block to where it already
is succeeds as a no-op **and records no undo snapshot**.

| Tracking | Behaviour |
|---|---|
| Off (`Accept`) | Detaches and re-inserts the EXISTING element. Descendant Unids, hyperlink/image/comment/note relationship ids all stay valid — v1 only moves within one part. |
| `RenderInline` | Keeps a revision-marked source and adds a revision-marked destination. A paragraph-like block uses one named `w:moveFrom`/`w:moveTo` pair (`ListRevisions` reports it as a single `move` revision that resolves both sides); a table lowers to Word's row delete + insert, which Word may present as two revisions. `EditResult.Created` carries the destination anchor so focus can follow the moved block. |

`accept ≡ the requested order` and `reject ≡ the original order` hold for both shapes.

**Two live copies, so identities must stay unambiguous.** A tracked move duplicates the block:

- bookmarks — rejected with `UnsupportedInlineBoundary` when the source contains bookmark
  markers. Both revision sides are live, so duplicating the name violates the first-class global
  name contract while keeping the markers on only one side loses them on accept or reject;
- comments — the move SOURCE takes a fresh comment id and a cloned definition (fresh
  `w14:paraId`, entries in both threading parts, cloned replies re-pointed at cloned parents),
  leaving the destination on the original comment and its thread;
- footnote/endnote references — deliberately duplicated: a note cited at both the old and new
  position is a faithful pending move, and exactly one citation survives either resolution.

### Refusals

`InvalidPosition` for: a different parent or part; a source owning an inline `w:pPr/w:sectPr`
(a section boundary, not an ordinary visual block); a move that would change or invert a
cross-block comment, bookmark, permission or native-move range; a move whose span crosses a
section-break paragraph; a source already inside a native move range; and — in tracked mode
only — a source that already contains revision markup a move would have to re-wrap.
Tracked moves containing bookmark markers are also refused with `UnsupportedInlineBoundary`.
`AnchorWrongKind` for a non-block kind or a non-top-level block (a table cell paragraph: move
the whole table instead).

### `ValidMoveTargets` — ask before offering

Returns the blocks `sourceAnchorId` may legally move next to, in document order; empty when the
block cannot move at all. It shares `MoveBlock`'s guards rather than restating them, so a listed
`(target, side)` pair is one `MoveBlock` accepts. That sharing has to reach the **mode-dependent**
rejections too: in `RenderInline` mode a bookmark-bearing source cannot move at all, and since
Word puts a `_Toc` bookmark on every heading, a rejection this gate did not know about would draw
drop indicators over most of a TOC'd contract before every drop hard-failed.

**The two sides are reported separately, and that distinction is load-bearing.** Landing *into* a
cross-block bookmark/comment/permission range changes its membership while landing outside it does
not, so a target is routinely legal `Before` and refused `After` (or the reverse) — a caller that
knows only "this target is reachable" will pick the refused side half the time. The editor gates
its drop indicator on the side, snaps a drop to the legal side when only one is, and resolves each
move-menu command against `(target, side)` pairs. Because a section break partitions the body,
"move to top/bottom" means the ends of the source's own region.

Note the practical consequence on real documents: a heavily cross-referenced contract can restrict
moves far more than its section breaks alone would suggest, because most blocks sit inside some
range. That is the confirmed v1 safety policy (reject rather than silently re-scope a range), not a
defect — and it is why the UI must ask rather than offer everything and fail on drop.

`MoveBlock` stays authoritative: the gate is advisory, and a caller that bypasses it still gets the
typed error. WASM/npm only, like the other editor-support endpoints.

**Cost.** One call answers a whole drag, so it is built that way: the facts that belong to the
CONTAINER — block order, which blocks own a section break (kept as a prefix sum), and each
cross-block range as a pair of block indices — are computed once, and each of the 2N candidate
questions is then index arithmetic. A move relocates exactly one element, so the reordered
sequence is a function of three indices and needs no second list; a range survives iff its two
endpoints still bound a window of the same width, which is equivalent to comparing the member
sets because the only element that can enter or leave is the source (including when the source IS
an endpoint, where any relocation but the identity one changes the width).

That equivalence is not obvious, so it is tested rather than argued: `DocxSessionMoveBlockTests`
carries the original set-membership implementation as an oracle and asserts the two agree for
every (source, target, side) triple over documents with nested, overlapping, table-spanning,
single-block, inverted, and dangling ranges plus section breaks. The naive version cost seconds
on a 234-block charter with 392 bookmarks, which the drag UI paid on every drag start.

## Bulk block removal — `DeleteRange` and `DeleteSection`

### `DeleteRange` — bulk sibling removal

`session.DeleteRange(fromAnchorId, toAnchorIdExclusive)` deletes every top-level
block-level sibling between two anchors. Both endpoints must:

- Be block-level kinds (`p`, `h`, `li`, `tbl`).
- Live in the same package part (same scope).
- Share a direct parent (the call refuses to span into nested containers like
  table cells; use a per-cell `DeleteBlock` loop for those).
- `from` must precede `to` in document order.

Records **one** undo snapshot — `Undo()` after `DeleteRange` restores every
removed element together. `EditResult.Removed` lists every anchor (including
descendant anchors of removed blocks) that disappeared.

**Tracked-change mode** (`Settings.TrackedChanges = RenderInline`): `DeleteRange`
wraps each removed paragraph's runs in `w:del` and marks the paragraph mark
itself as deleted via `w:pPr/w:rPr/w:del`. Tables get `w:trPr/w:del` on every
row (Word's row-deletion convention — there is no table-level "delete" markup),
plus the same run/paragraph-mark wrapping inside every cell. Nested tables
recurse.

Block-level `w:sdt` content controls are reversible too. Two paired
`w:customXmlDelRangeStart` / `w:customXmlDelRangeEnd` ranges cross the control's
opening and closing tags, matching Word's native content-control deletion shape.
Payload paragraphs and tables receive their normal deletion markup recursively;
nested block controls receive their own paired ranges. Accepting the revisions
therefore removes the control and its payload, while rejecting restores the
original wrapper, metadata, and content. Locked (`w:lock`) and data-bound
(`w:dataBinding`) controls use the same shape—the lock and binding metadata are
preserved until the revision is resolved.

Anchor accounting describes what actually happened. Ordinary paragraph/table
top-level anchors remain the compact `Modified` contract. A structured wrapper
has no anchor of its own, so every anchored descendant retained under that
wrapper appears in `Modified`, without duplicates. A remaining structural
fall-through that must be hard-removed appears in `Removed`; it is never silently
omitted from both lists.

`w:customXml` wrappers are deliberately unsupported in tracked bulk deletion.
If any selected block contains one, the operation fails before taking an undo
snapshot or changing the document with `IncompatibleElementType` and a message
identifying `w:customXml`. This is the explicit unsupported branch of the
custom-XML deletion contract; accepted-mode bulk deletion remains unchanged.

### `DeleteSection` — heading-bounded bulk removal

`session.DeleteSection(headingAnchorId)` deletes a heading and every sibling
below it up to (but not including) the next heading at the same or higher
level. "Level" matches the projection's notion: `Heading1` = 1, `Heading2` = 2,
…, `Title` = 1, `Subtitle` = 2.

If the target heading has no sibling-heading boundary after it, the section
extends to the end of the parent.

Built on `DeleteRange` semantics via the shared `DeleteSiblingRangeCore` helper:
same undo, same `EditResult` accounting, the same native `w:sdt` envelope and
recursive payload markup, the same pre-mutation `w:customXml` refusal, and the
same reported structural fall-through.

## Native hyperlinks and bookmarks

Hyperlinks and bookmarks are addressable document objects, not projection-only formatting:

```csharp
IReadOnlyList<HyperlinkInfo> ListHyperlinks(ProjectionScopes scopes = ProjectionScopes.All);
EditResult AddHyperlink(string anchorId, CharSpan span, HyperlinkTarget target);
EditResult UpdateHyperlink(string hyperlinkId, HyperlinkTarget target);
EditResult RemoveHyperlink(string hyperlinkId);

IReadOnlyList<BookmarkInfo> ListBookmarks(ProjectionScopes scopes = ProjectionScopes.All);
EditResult AddBookmark(string name, DocumentRange range);
EditResult RenameBookmark(string name, string newName);
EditResult MoveBookmark(string name, DocumentRange range);
EditResult RemoveBookmark(string name);
```

`HyperlinkTarget.External(uri)` creates or reuses a hyperlink relationship on the XML part that
owns the link. A header link is related from its `HeaderPart`, a footnote link from its
`FootnotesPart`, and so on; the main document never acts as a relationship proxy for another
story. `HyperlinkTarget.Internal(bookmarkName)` writes only `w:anchor` and requires exactly one
coherent, ordered start/end pair with that globally unique name in one story part; a lone or
ambiguous marker is `MissingBookmarkTarget`, not a targetable bookmark. Wire callers must pass
target kind `internal` or `external`; unknown strings are `InvalidHyperlinkTarget` rather than
silently becoming external. Orphaned external relationships are removed only after their last markup
reference disappears. The owner-aware relationship helper is generic over the referencing
attribute so part-backed content such as images can reuse the same ownership/orphan rules.

`HyperlinkInfo` reports the owner part/scope, enclosing anchor, exact half-open `CharSpan`, visible
text, target, relationship metadata, and broken-target state. A hyperlink id follows the anchor
identity contract: stable for the live session and across `Save(true)` / `PersistAnchorIds`, but
not promised across a default save that strips Docxodus Unids. Splitting a paragraph inside a link
cuts the `w:hyperlink` in two and gives each half its own identity, so both stay individually
addressable by `UpdateHyperlink`/`RemoveHyperlink`.

`AddHyperlink` relocates the whole contiguous sibling range it covers, not just the `w:r` elements.
A zero-width marker sitting between two selected runs — `w:bookmarkStart`/`End`,
`w:commentRangeStart`/`End`, `w:proofErr` — is a legal `w:hyperlink` child and moves *inside* the
new link at its original position. Stranding it after the link would, for a bookmark whose start
lies inside the span, put the start after its own end and break the pair permanently. Content
outside the span keeps its side, so document order is preserved for every marker.

Story scope covers body, headers, footers, footnotes, endnotes, **and comments**: a comment
paragraph is anchor-addressable (`p:cmt:…`) and editable, so it owns its own hyperlink
relationships and can host bookmark markers like any other story.

Bookmark names follow Word's UI-safe form: 1–40 characters, starting with a letter or underscore,
then letters, digits, or underscores. **Word's own namespace is closed to creation.** `AddBookmark`
and `RenameBookmark`'s destination name refuse `_GoBack`, `_Toc*`, `_Ref*`, `_Hlt*`, and `_Hlk*`
with `InvalidBookmarkName`: Word allocates and rewrites those for itself (a TOC refresh regenerates
the whole `_Toc*` family), so a name placed there is reallocated or clobbered under the caller.
Bookmarks Word already put there are *not* frozen — they list, rename, move, and remove like any
other, with their cross-reference fields retargeted or blocking removal as usual. Renaming a `_Toc*`
bookmark is simply not durable, because the next TOC refresh regenerates it.

Names are globally unique; numeric `w:id` pairing is scoped
to the owning story part because real Word files reuse numeric ids across parts. A
`DocumentRange` may cross paragraphs but both endpoints must belong to the same body, individual
header/footer, footnote, endnote, or comments part. `BookmarkInfo.Range` carries the two endpoint anchors and
offsets; `Segments` supplies exact per-paragraph spans and text. Unmatched starts and ambiguous
same-story numeric ids or duplicate names remain visible as invalid diagnostics. Orphan end markers
have no name/start coordinate and are not returned as rows, but still participate in fresh numeric
id allocation. `_Docxodus_Ann_*` bookmarks are owned by the annotation subsystem and reject generic
bookmark mutation.

**A bookmark has two consumer families, and both count as "referenced."** Beyond
`w:hyperlink/@w:anchor`, Word cross-references a bookmark through `REF`, `PAGEREF`, `NOTEREF`, and
`HYPERLINK \l` field instructions — carried either in `w:fldSimple/@w:instr` or in the
`w:instrText` runs between a `w:fldChar` begin and its matching separate/end, and split across
several runs as often as not. Every entry in a Word table of contents is a `PAGEREF` over a `_Toc`
bookmark. Rename first requires one coherent same-story pair, then changes the start marker and
retargets **both** families across all stories in one undo step, splicing only the name token so
switches (`\h`, `\* MERGEFORMAT`) survive verbatim; a split instruction is coalesced onto its first
`w:instrText`. Seeing only the anchor links would leave "Error! Bookmark not defined." behind.
`BookmarkInUse` is judged against both families too, so removing a bookmark a TOC still cites is
refused rather than silently dangling it.

**Know the reach of that guard before you rely on it.** It also gates structural deletion, and it is
reference-scoped: a deletion is refused only while a reference *survives outside* the deleted
region. Word puts a `_Toc` bookmark on every heading and keeps the matching `PAGEREF` in a different
paragraph, so on a TOC'd document `DeleteBlock(heading)` is refused — in **default, untracked** mode,
with no force/opt-out. Delete the citing field (or the whole TOC) first, or rename rather than
remove. This is the same no-dangling-reference policy that has always covered `w:anchor` links; what
changed is how much of a real contract it reaches.

Move validates the destination before detaching the old pair. It retains its numeric id for a
same-part move; a **cross-part** move takes a fresh document-global id, because `w:id` is
part-scoped and carrying it into a part that already uses that decimal makes *both* bookmarks
unresolvable. Structural edits likewise reject a pair crossing the deletion
boundary, a targeted pair, or a managed pair before snapshotting. Whole-paragraph replacement keeps
endpoint character coordinates and clamps them to the new end; surgical replacement, split, merge,
and direct block moves preserve marker order. First-class hyperlink/bookmark metadata mutations are
explicitly unavailable in `RenderInline` mode (`TrackedOperationUnsupported`), because Word has no
faithful native revision shape for them. For the same reason, tracked whole-paragraph replacement
rejects a paragraph containing bookmark markers before snapshotting; tracked surgical span
replacement remains supported because it keeps zero-width markers in place.

All methods route through `DocxSessionOps` and are surfaced by WASM/npm, stdio/`docx-scalpel`, and
MCP's `docxodus_links`. Markdown `[text](uri)` and `[text](#bookmark)` use the same target validation,
part ownership, relationship reuse, and cleanup rules.

## Finding anchors via tagged annotations

The session addresses content by anchor id, but real workflows don't start with anchor ids — they start with intent ("edit the indemnification provision," "tighten the termination clause"). The clean way to bridge intent to anchors is to **annotate the regions ahead of time**, then resolve the annotation to its anchor(s) at edit time.

Docxodus's `AnnotationManager` already persists annotations into the docx itself: each annotation creates a `w:bookmark` named `_Docxodus_Ann_<id>` covering the range, and a custom XML part stores the metadata (`LabelId`, `Label`, `Color`, `Metadata` key/value bag). See [`custom_annotations.md`](custom_annotations.md) for the full mechanism and lifecycle. Annotations survive save/reopen and travel with the document.

`DocxSession` exposes four discovery helpers that read directly off the long-lived `WordprocessingDocument` (no save/reopen round-trip per call):

```csharp
session.ListAnnotations();                          // every annotation in the doc — id, labelId, label, color, author, annotatedText
session.FindByAnnotation("ann-id");                 // IReadOnlyList<AnchorTarget> — the blocks the bookmark covers
session.FindByLabel("INDEMNIFICATION");             // IReadOnlyDictionary<annotationId, IReadOnlyList<AnchorTarget>>
session.FindByBookmark("_Docxodus_Ann_ann-id");     // lower-level: resolve any bookmark name (managed or user-authored)
```

The canonical agentic recipe collapses to:

```csharp
foreach (var (id, anchors) in session.FindByLabel("INDEMNIFICATION"))
    foreach (var a in anchors.Where(a => a.Anchor.Kind is "p" or "h" or "li"))
        session.ReplaceText(a.Anchor.Id, "Revised indemnification language…");
```

What `FindByAnnotation` / `FindByLabel` / `FindByBookmark` return in v1:

- **All block-level anchors whose subtree overlaps the bookmark range, in document order, deduplicated.** That includes the immediate paragraph plus any enclosing table / row / cell, so an agent sees "this annotation lives in a table" without re-walking the tree. Filter by `Anchor.Kind in {"p","h","li"}` when you want only the text-bearing blocks suitable for `ReplaceText`.
- **Empty list when the id/label/bookmark is unknown** or the bookmark's end marker is missing. No exceptions for not-found.
- **All story scopes for generic bookmarks.** `FindByBookmark` resolves body, header, footer,
  footnote, and endnote bookmarks. `AnnotationManager` itself still authors managed annotation
  bookmarks in the main document part.

Two addressing details matter:

- **`FindByBookmark` returns enclosing block anchors**, for compatibility with annotation-driven
  workflows. Use `ListBookmarks` when exact endpoint ranges and per-paragraph character spans are
  required.
- **Marker-preserving edits are deterministic.** Whole replacements retain/clamp endpoint offsets;
  surgical replacements, paragraph split/merge, and direct block moves preserve marker order.
  Structural edits that would orphan a marker fail rather than silently corrupting the range.

The agent's prompt should also be aware: it can call `session.ListAnnotations()` once at session start to enumerate available labels (e.g., "you can target: INDEMNIFICATION, TERMINATION, GOVERNING_LAW") and present those as tools rather than asking the LLM to discover them from text.

## Headers, footers, and page-number fields

Running headers/footers and page-number fields live in their own OOXML parts
(`HeaderPart`/`FooterPart`), *outside* the body — so before issue #236 the session
could only *inspect* them (`GetSectionInfo` → `HeaderPartUris`/`FooterPartUris`),
never author them. These three methods close that gap; they're exposed in .NET,
WASM/npm, and stdio/`docx-scalpel`.

### Methods

| Method | What it does |
|---|---|
| `SetHeaderText(anchorId, HeaderFooterKind, markdown)` | Set the running header story for the section that owns `anchorId`. |
| `SetFooterText(anchorId, HeaderFooterKind, markdown)` | Same, for the footer. |
| `InsertPageNumberField(anchorId, PageNumberField = CurrentPage)` | Append a `PAGE`/`NUMPAGES` field to a paragraph (usually a header/footer one). |
| `EnsureHeaderFooterVisible(anchorId, HeaderFooterKind)` | Make the section's first/even stories actually render (`w:titlePg` / `w:evenAndOddHeaders`). |

**Why `EnsureHeaderFooterVisible` exists.** `SetHeaderText`/`SetFooterText` set the visibility
flags *while writing content*, which covers authoring a story from scratch. It does not cover a
document that already carries a first/even reference with the flag absent — and that is exactly
what Word leaves behind when "Different first page" / "Different odd & even pages" is switched
back off: the part and its reference stay, only the flag goes. Editing such a pre-existing story
through the anchor-addressed text ops then yields a file whose header content is present but
never rendered. The flags belong to the **section**, not to a content write, so this is its own
operation: `First` → `w:titlePg`, `Even` → the document-global `w:evenAndOddHeaders`, `Default`
→ a successful no-op. Idempotent; a non-body anchor is `AnchorWrongKind`.

`TestFiles/HC031-Complicated-Document.docx` is the canonical example — all six stories present,
neither flag set (`DS268`).

**Addressing.** `SetHeaderText`/`SetFooterText` take *any body block* in the target
section — the governing `w:sectPr` is resolved the same way `GetSectionInfo` resolves
it (a forward mid-document section break, else the body's trailing `sectPr`; if the
body has none, a trailing `sectPr` is synthesized). This mirrors `GetSectionInfo`
returning `null` for non-body anchors: passing a header/footer/footnote anchor is an
`AnchorWrongKind` error, because a story attaches to a *body* section.

**`HeaderFooterKind`** = `Default` / `First` / `Even`, mapping to the reference's
`w:type`. `First` additionally sets the section's `w:titlePg`; `Even` sets
`w:evenAndOddHeaders` in the settings part — without those flags Word ignores the
first/even story. Calling the same kind twice **reuses** the existing part and
replaces its content (so `SetFooterText` is an idempotent "set the footer to this");
a different kind creates a second part/reference.

Two `Even` sharp edges worth knowing:

- `w:evenAndOddHeaders` is **document-global and governs footers too**. Once set,
  even pages stop inheriting the Default stories entirely — a section with only a
  Default footer shows *no footer at all* on even pages (spec-correct Word behavior,
  observed identically in LibreOffice). If you set an Even header and want footers to
  keep appearing on every page, set an Even footer too.
- The flag is inserted at its CT_Settings schema slot via
  `WordprocessingMLUtil.EnsureEvenAndOddHeaders` (shared with the DocxDiff
  header/footer renderer); every other settings child — including ones the ordering
  table doesn't know, like `w:hdrShapeDefaults`/`w:shapeDefaults` that real Word
  documents carry — stays exactly where it was. (An earlier whole-part reorder
  corrupted such documents; `DS263`/`DS264` pin the fix.)

**Content & styling.** The `markdown` uses the same subset as `InsertParagraph`
(bold/italic/links/etc.). Each paragraph with no explicit style gets the built-in
`Header`/`Footer` paragraph style, so it inherits Word's centre-of-page and
right-margin tab stops — the layout page-number footers rely on. An empty payload
yields one empty story paragraph.

**Return shape.** `SetHeaderText`/`SetFooterText` report the new story paragraphs in
`EditResult.Created` (scope `hdr{N}`/`ftr{N}`, 1-based by part-collection order) — pass
one to `InsertPageNumberField`. `InsertPageNumberField` reports the target paragraph in
`Modified`. The field is a native complex field (`fldChar`/`instrText`, cached "1"
result), so it renders and updates like a hand-authored field.

### Recipe — the S-1 running footer ("Last Updated … / centered page N")

```csharp
var body = session.Project().AnchorIndex.Values
    .First(t => t.Anchor.Kind == "p" && t.Anchor.Scope == "body").Anchor.Id;

var footer = session.SetFooterText(body, HeaderFooterKind.Default, "Last Updated October 2025");
var footerPara = footer.Created[0].Id;                       // scope "ftr1"
session.SetParagraphFormat(footerPara, new ParagraphFormatOp { Alignment = ParagraphAlignment.Center });
session.InsertPageNumberField(footerPara, PageNumberField.CurrentPage);
```

### Undo/redo and the snapshot reconcile

`SetHeaderText`/`SetFooterText` can *add* an OOXML part, which the session's per-part
snapshot didn't previously track (only the annotations custom-XML part was
create/delete-reconciled). `DocumentSnapshot` now also records each header/footer
part's relationship id, and `RestoreSnapshot` reconciles them: on undo it deletes
parts the snapshot lacks; on redo it re-creates the ones it has **with their original
relationship id** (via `AddNewPart<HeaderPart>(relId)`) so the just-restored `sectPr`
reference resolves. Content of surviving parts restores by URI as before. One edge is
documented as intentional: the `w:evenAndOddHeaders` settings flag (only set by the
`Even` kind) isn't reverted by undo — it's idempotent and has no visual effect without
an even story.

### Which part supplies which kind — `SectionInfo.HeaderRefs`/`FooterRefs`

`HeaderPartUris`/`FooterPartUris` report *which* parts a section references but not which
story kind each one supplies, and the projection's `hdr{N}`/`ftr{N}` numbering is by
part-collection order, which carries no kind information either. A client that wants to
show or edit "this document's **first-page** header" therefore cannot resolve it from
those lists.

`SectionInfo.HeaderRefs` and `FooterRefs` close that: each entry is a
`HeaderFooterRef { HeaderFooterKind Kind; string PartUri; bool Inherited; }`, in the
references' declaration order. `w:type` is optional in OOXML, so an absent (or
unrecognized) value reads as `Default` per ECMA-376 §17.6.10.

**They report the stories that EFFECTIVELY apply, not just the section's own.** A section
that declares no reference of a given type *continues the previous section's*
(ECMA-376 §17.6.17), which is why a multi-section document typically defines its headers
once in the first section and leaves the rest empty — `HC031-Complicated-Document.docx`
has four sections and only the first declares anything. Reporting own references alone
would tell a caller "this part of the document has no header" when it visibly does, and an
editor acting on that would mint a redundant part and break the inheritance. Inherited
entries carry `Inherited = true`; editing one edits the part both sections share, which is
what Word does. `HeaderPartUris`/`FooterPartUris` keep their original meaning — this
section's **own** references only.

Combined with the `partUri` each projection anchor already carries, this gives a client
the full kind → part → story-paragraph-anchors chain:

```csharp
var info = session.GetSectionInfo(body)!;
var firstHeaderPart = info.HeaderRefs.First(r => r.Kind == HeaderFooterKind.First).PartUri;
var storyParas = session.Project().AnchorIndex.Values
    .Where(t => t.PartUri == firstHeaderPart && t.Anchor.Kind == "p")
    .Select(t => t.Anchor.Id);
```

Wire: `headerRefs`/`footerRefs` (npm `SectionInfo.headerRefs`/`footerRefs`, `HeaderFooterRef`;
`docx-scalpel` `SectionInfo.header_refs`/`footer_refs`). This is exactly what the browser
editor's band chrome uses to label its kind selector.

### Page-number formatting (issue #277)

Page numbering has **two independent layers**, and conflating them is the usual way to get
it wrong.

**The section** — `SetPageNumbering(bodyAnchor, PageNumberingOp)` writes `w:pgNumType`,
which is exactly what Word's *Format Page Numbers…* dialog writes. `Start` (`w:start`) is
the number the section begins at; `Format` (`w:fmt`) is the format its pages use. Both
fields are tri-state — null leaves that **attribute** alone, so the start can be set
without disturbing the format and vice versa. `ClearPageNumbering(bodyAnchor)` removes the
two attributes (preserving the chapter-numbering ones this surface never writes, and
removing the element only once nothing is left on it). Addressed by any body block, with
the governing `w:sectPr` resolved exactly as `GetSectionInfo` resolves it, and synthesized
if the body has none.

This is the normal way to number pages: set the section once, insert plain fields. A
`PAGE` field with no switch renders through it.

**The field** — `InsertPageNumberField(anchor, field, format?)` writes the field's own
`\*` general-formatting switch (`PAGE \* roman`). Omitting `format` — the default —
emits a plain field, byte-for-byte what earlier versions emitted. A switch *overrides* the
section for that one field and keeps overriding it if the section later changes, so it is
the escape hatch, not the default route. The editor's band deliberately inserts plain
fields for exactly this reason.

Both reject `NumberFormat.Bullet` and a negative start with
`EditErrorCode.InvalidPageNumbering`: a bullet is a valid **list** format with no
page-number counterpart in either vocabulary, so accepting it could only mean silently
writing something else.

`NumberFormat` is reused rather than duplicated — it is already this library's name for
`ST_NumberFormat`, which is the type of both `w:numFmt` and `w:pgNumType/@w:fmt`.
`Docxodus/Internal/NumberFormats.cs` is the single owner of the three mappings (OOXML
token, `\*` switch argument, rendered glyph); the switch spellings are case-significant,
since `roman` is `i, ii, iii` and `ROMAN` is `I, II, III`. A field's cached result is
seeded with page 1 rendered in the requested format (`i`, `A`, `1`) rather than a
hardcoded `"1"`, so a renderer that does not recompute fields agrees with the switch.

`w:pgNumType` has a fixed slot in the `CT_SectPr` sequence. `WordprocessingMLUtil`'s
`Order_sectPr` + `InsertSectPrChildInOrder` place it there — the same slot-insert
discipline as `EnsureSettingsChildInOrder`, and now the single owner of that ordering for
`w:titlePg` too (`DocxSession` and `IrMarkupRenderer` previously each carried a private
"what follows titlePg" list).

**Read-back:** `SectionInfo.PageNumberStart` / `PageNumberFormat` (wire
`pageNumberStart`/`pageNumberFormat`; `docx-scalpel` `page_number_start`/
`page_number_format`). Both are *omitted* when the attribute is absent rather than
defaulted, because "continues the previous section in Word's default format" is a
different claim from "starts at 1 in decimal" — a UI that cannot tell them apart writes
attributes the document never had.

### The editor region

The browser `DocxEditor` ships the visual affordance as **docked bands** — see
`docs/architecture/ir_editor_feasibility.md` § "Header/footer editing region". Story
paragraphs there are ordinary editable blocks addressed by their `p:hdr1:<unid>` anchors,
which every text/format mutation on this page already accepts.

The band chrome carries the section's page-number **format** and **start-at** controls
(`setPageNumbering`/`clearPageNumbering`/`pageNumbering` on `DocxEditor`). They sit on both
bands and show the same values, because they describe the section rather than either story.

### Not yet

- **Deleting a header/footer story.** `SetHeaderText`/`SetFooterText` create or replace;
  there is no operation that removes a part and its reference. An empty payload yields an
  empty story, which is not the same thing.
- **The in-page-margin editing overlay** (editing the running head *inside* the page box's
  top margin in `{ paginated: true }` mode). Two things block it: the full-document render
  assigns Unids to the main document part only, so header/footer content carries no
  `data-anchor`; and pagination clones one header node onto every page, so N pages would
  mean N DOM nodes sharing one anchor. The docked bands avoid both.
- **Chapter page numbering** (`w:pgNumType/@w:chapStyle`/`@w:chapSep` — "1-1, 1-2" numbers
  derived from a heading style). `ClearPageNumbering` preserves those attributes when a
  document already has them, but nothing writes them.
- **Recomputing a field's cached result.** The number a page-number field shows in a
  non-paginated render is the value cached in the file; Word recomputes on open, and the
  paginated preview substitutes the real per-page number (below), but a continuous-mode
  render still shows the cached one.

## Tier B: footnotes & endnotes (issue #276)

The projection has always *read* notes — the `fn`/`en` scopes, `EditSummary.FootnoteCount`,
`ReplaceText` on a note's paragraph, `DeleteBlock` on a note definition (which also strips
every reference to it). What was missing was the verb that *creates* one. `InsertFootnote` /
`InsertEndnote` close that; there is no separate "edit note" or "delete note" op because the
existing anchor-addressed ops already are those.

### Methods

| Method | What it does |
|---|---|
| `InsertFootnote(anchorId, characterOffset, markdown)` | Create a footnote with body `markdown` and cite it from body paragraph `anchorId`, `characterOffset` characters into its text. |
| `InsertEndnote(anchorId, characterOffset, markdown)` | Same, into the endnotes part, emitting a `w:endnoteReference`. |

**Addressing.** A **body** paragraph/heading/list-item anchor, plus a character offset in
`[0, len(paragraph text)]` — `0` places the citation before all text, `len` after all of it.
An out-of-range offset is `OffsetOutOfRange`. Non-body anchors are `AnchorWrongKind`: Word does
not allow a note reference inside a header/footer story or inside another note, and the `fn`/`en`
scopes are note *definitions*, not citation hosts. Rejecting is deliberate — the alternative is
emitting a document Word offers to repair.

The offset is resolved through the same `SplitRunsAtOffset` + `SplitInlineContainersAtOffset`
pair that `SplitParagraph` and `ApplyFormat` use, so a citation lands cleanly mid-run and inside
a hyperlink without a second offset walker to drift.

**What the first note in a document creates.** On a package with no `FootnotesPart` (or
`EndnotesPart`), the op writes the whole scaffold Word writes, not just the definition:

- the part itself, holding the two notes Word reserves for page-rendering scaffolding —
  `w:type="separator"` at id `-1` and `w:type="continuationSeparator"` at id `0`. The
  projector already filters these out of the anchor index (`IsBoilerplateNote`), and
  `DeleteBlock` already refuses to remove them.
- the `w:footnotePr`/`w:endnotePr` settings declaration naming those two ids, inserted at
  its CT_Settings schema slot via the shared `EnsureSettingsChildInOrder` — the settings part
  is never wholesale reordered (same discipline as `w:evenAndOddHeaders`; see the header/footer
  section for the document this protects).
- the `FootnoteText`/`EndnoteText` paragraph style and the superscript
  `FootnoteReference`/`EndnoteReference` character style, find-or-create via `StyleFactory`, so
  a house style already in the document wins and a document that had neither doesn't render the
  citation as full-size body text.

A second note reuses all of it and only appends a definition.

**Note id allocation — ids must ascend in REFERENCE order.** This is an invariant every
Word-authored document holds (verified across the `TestFiles` corpus — including documents whose
ids have gaps, e.g. 17/21/26 — and a 94-footnote real-world model certificate). Renderers depend
on it: LibreOffice numbers the body markers by citation position but pairs them against the
**id-sorted** definition list, so a first-cited note holding the highest id renders the *wrong
note text* — the marker reads "1" and points at somebody else's footnote. Nothing errors; the
document is simply, silently wrong.

So the allocator works in two cases:

- **Citation follows every existing one** (the common case): take one above the highest id used by
  any definition in the note part *or* any reference anywhere in the package (body, headers,
  footers, both note parts). Appending keeps ids ascending, so nothing else moves. Scanning
  references too — not just definitions — is what stops a document with gaps from aliasing an
  existing definition. `DS322` pins this with a fixture whose user notes are ids 1, 5 and 9.
- **Citation lands before an existing one**: the new note takes the smallest id cited after it, and
  every user note at or above that id shifts up by one — definitions *and* every reference to them
  in every part. Notes cited earlier keep their ids. Word-reserved notes (any `w:type`:
  `separator`, `continuationSeparator`, `continuationNotice`) are never renumbered; their ids sit
  below every user id, so shifting upward cannot collide. Taking the *minimum* of the following ids
  rather than the first keeps this correct even on an input document that already violated the
  invariant. `DS336`/`DS337` pin the ordering and that each shifted note keeps its own text.

Because a shift rewrites `w:id` on renumbered definitions, and the note-definition anchor's unid is
derived from that id, the shifted notes' `fn`/`en` anchors change. Their paragraph anchors — and
every body anchor — are unaffected.

**Markup.** Word-faithful on both sides:

```xml
<!-- body -->
<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r>

<!-- footnotes.xml -->
<w:footnote w:id="1">
  <w:p>
    <w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>
    <w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>
    <w:r><w:t xml:space="preserve"> </w:t></w:r>
    <w:r><w:t>Source: 2025 annual report.</w:t></w:r>
  </w:p>
</w:footnote>
```

The `w:footnoteRef`/`w:endnoteRef` auto-number mark plus its separating space go on the first
paragraph of the note only, so the note reads "1 Source: …" rather than "1Source: …".
The body is the same markdown subset as `InsertParagraph`; an empty payload yields one empty
note paragraph.

**Return shape.** `Created` carries the note **definition** anchor (kind `fn`/`en`) followed by
its paragraph anchors (kind `p`, scope `fn`/`en`); `Modified` carries the citing body paragraph.
That is the whole lifecycle handoff:

```csharp
var made = session.InsertFootnote(bodyAnchor, 0, "Source: 2025 annual report.");
var notePara = made.Created.First(a => a.Kind == "p" && a.Scope == "fn").Id;
var noteDef  = made.Created.First(a => a.Kind == "fn").Id;

session.ReplaceText(notePara, "Source: 2026 annual report.");  // edit the note
session.DeleteBlock(noteDef);                                  // drop it + the body citation
```

### Undo/redo and the snapshot reconcile

Creating the first note *adds an OOXML part*, the same problem `SetHeaderText`/`SetFooterText`
solved. `DocumentSnapshot` therefore records the footnotes/endnotes parts' relationship ids
alongside the header/footer ones, and `RestoreSnapshot` runs a `ReconcileNoteParts` twin of
`ReconcileHeaderFooterParts`: undo deletes a part the snapshot lacks, redo re-creates it with its
original relationship id (`AddNewPart<FootnotesPart>(relId)`). Content of surviving parts restores
by URI as before. `DS328` pins undo-removes-part / redo-restores-part.

### `FootnoteRefNotSupported` — narrowed, not retired

A `[^label]` reference inside a **markdown payload** is still rejected. The error code stays
(it is public surface and clients switch on it); only its meaning narrowed: a bare label can't
be resolved to a note the payload doesn't define, so the message now names the dedicated op.
Authoring a note is `InsertFootnote`/`InsertEndnote`; the payload subset is for a note's *body*,
not for minting citations.

### Rendering notes in the editor

An editor that can author a footnote has to be able to *show* it. The browser `DocxEditor` renders
notes as ordinary editable content, which took three things:

- **The render profile emits notes.** `DocxSessionOps.RenderHtml` and the editor's first-paint
  `completeArgs` both set `RenderFootnotesAndEndnotes`. They must stay in step — the remount output
  is required to match the first paint byte-for-byte. Footnotes are document *content*, not an
  editing affordance, so unlike the header/footer bands this is not opt-in: a document that has
  notes shows them.
- **Note paragraphs are anchor-stamped.** `HtmlConversionOps.AssignAnchorUnids` assigns the
  deterministic Unids to the footnotes/endnotes parts as well as the main part, so note paragraphs
  carry `data-anchor` and the editor wires them as ordinary blocks — no new command code, the whole
  ribbon works inside a note. `FindByUnid` searches the note parts too, so the stateless
  `RenderBlockHtml` can re-render a single note after an edit. Header/footer parts are deliberately
  *not* stamped: paginated output clones one header node onto every page, so a stamped header anchor
  would exist N times in the DOM. Each note renders exactly once, so notes have no such problem.
- **The citation marker is inert chrome.** A citation is a zero-width `w:footnoteReference`; the
  displayed number is computed by the renderer from document order, and the note backref (`↩`) is
  generated too. None of it is in the session's run text, so the editor excludes all of it from its
  content-offset space via one `GENERATED_CHROME_SELECTOR` (shared with generated list markers).
  Each place that has to honour it fails differently if missed: offsets drift (`OffsetOutOfRange`,
  silently dropped edits); the display number gets **committed as literal text**, destroying the
  citation run; or the user deletes a marker outright and orphans its note.

**Both editor modes show notes, differently.** Continuous mode renders the converter's
`<section class="footnotes">` at the end of the body flow. Paginated mode is richer: `pagination.ts`
already had a footnote engine (per-page distribution, continuations, splitting), and turning the
render flag on activates it — notes land in a note area at the bottom of the page that cites them,
above a separator rule, and a note too long for its page continues onto the next. Endnotes render as
a `section.endnotes` appended after the page stack rather than on their own final page, which is a
layout imperfection, not a correctness one: they are visible and editable. A note split across pages
puts a *different paragraph* of the note in each half, so each half stays independently addressable
— no two editable nodes ever share one anchor.

Note content lives inside the body flow, so anything that walks "the body blocks" must exclude
`section.footnotes`, `section.endnotes` and `.footnote-item` — the header/footer band already does
the equivalent by ignoring an anchor `GetSectionInfo` can't resolve to a body section (focusing a
note leaves the bands on the last body section rather than blanking them).

**Paginated mode's footnote engine.** `pagination.ts` already had per-page distribution, splitting
and continuations; turning the render flag on exercised it against a dense real document for the
first time and it needed four fixes, each worth knowing about because each failed *silently*:

- an unfitted note was held in a **single** continuation slot assigned once per note in a page's
  citation list, so a second unfitted note on the same page overwrote the first and it rendered
  nowhere — notes now queue and are *merged* into the next page's list, never replaced;
- the stylesheet's `>` combinators were XML-escaped (the CSS is the value of an `h:style` element),
  so the rule keeping a note's number and first line together was dropped by every browser;
- a note that couldn't be split was re-wrapped inside its own content, nesting a complete
  `.footnote-item` in another's content span;
- the content area spanned the full text height while the note block grew upward into it, so body
  and note glyphs could be painted on top of each other. The content area is now shrunk to what the
  notes leave, making that impossible by construction — worst case is a clean clip — and note
  heights are measured in the `.page-footnotes` styling context they render in so the reserve
  matches the paint.

`DS340`–`DS345` pin the engine half (marker + section emitted, note paragraphs stamped, stamped
anchors resolve through the session, edit-then-re-render, the stateless path, and reserved-note
filtering); `npm/tests/editor-footnotes.spec.ts` pins the browser half.

**Word-reserved notes never surface as editable blocks.** `IsBoilerplateNote` now treats *any*
typed note as reserved: ECMA-376 §17.11.17 defines the type as `normal` | `separator` |
`continuationSeparator` | `continuationNotice`, and only `normal` — which Word omits rather than
writes — is user content. Enumerating just separator/continuationSeparator let `continuationNotice`
through (real documents carry one), where it projected as a user note and, once the editor started
rendering notes, appeared as a stray empty footnote with no citation.

### Not yet

- **Moving a citation** without deleting and re-creating the note.
- **Note numbering options** (`w:numFmt`, restart-per-page/section, custom marks) — the
  `w:footnotePr` written on part creation declares only the separators; everything else is
  Word's default (`1, 2, 3…`, continuous).
- **Citing an existing note twice.** Each call creates a new definition; there is no
  "reference note N again" op.
- **Tracked-changes mode.** `Settings.TrackedChanges = RenderInline` does not wrap the citation
  in `w:ins` — consistent with every other insert op (`InsertParagraph`, `InsertTable`,
  `InsertHorizontalRule`, `SetHeaderText`); only `ReplaceText`/`DeleteBlock`/`DeleteRange`/`DeleteSection` track.
- **Narrowed projection scopes.** A session opened with `ProjectionSettings.Scopes` excluding
  `Footnotes`/`Endnotes` still writes the note correctly, but `Created` comes back without the
  note anchors — they resolve against a projection that omits the part. Family behavior, identical
  to `SetHeaderText` with `Headers` excluded.

## Comments (issues #300, #317, and #341)

Native Word comment authoring — real `w:comment` markup the Reviewing pane shows, not the
Tier E annotation overlay (which stays: it solves a different problem, semantic tagging for
external tools). Follows the #274/#276 part-creation pattern: the `WordprocessingCommentsPart`
and the `CommentText`/`CommentReference` styles are find-or-created on first use, and part
create/delete is undo/redo-reconciled by `ReconcileCommentsPart` (the `ReconcileNoteParts`
twin — `DocumentSnapshot` carries the part's relationship id). Mechanics live in
`Internal/CommentOps.cs` (the `AnnotationOps` split). Exposed in .NET, WASM/npm
(`addComment`/`addCommentToRevision`/`addCommentReply`/`setCommentResolved`/`updateComment`/`removeComment`/
`listComments`), stdio/`docx-scalpel` (`add_comment`/`add_comment_to_revision`/
`add_comment_reply`/`set_comment_resolved`/`update_comment`/`remove_comment`/`list_comments`), and the MCP
server's `docxodus_comment` tool.

### Methods

| Method | Description |
|--------|-------------|
| `AddComment(anchorId, span?, author, markdown, initials?, date?)` | Comment on a character span of a body paragraph (`null` span = whole block; the same `SplitRunsForSpan` mechanism annotations use brackets the range with `w:commentRangeStart`/`End`, then the `CommentReference`-styled `w:commentReference` run lands directly after the rangeEnd). The definition gets Word's shape: `CommentText` paragraphs, a leading `w:annotationRef` mark run, `w:id` allocated max+1 over definitions *and* dangling markers. `w:date` is written only when provided — deterministic by default; an Unspecified-kind `DateTime` is treated as UTC. |
| `AddCommentToRevision(revisionId, author, markdown, initials?, date?)` | Comment on the exact live markup extent identified by `ListRevisions()`: insertion/deletion wrappers, the destination of a move, affected runs for run-format changes, or the affected paragraph for structural/property changes. Markers are placed outside revision wrappers, so accept/reject keeps the comment and either preserves its range on surviving content or collapses it to a point. Unknown or already-resolved ids return `RevisionNotFound`. |
| `AddCommentReply(parentCmtAnchorId, author, markdown, initials?, date?)` | Add a Word-native reply. It gets its own definition, `w:id`, and adjacent `w:commentReference`; range start/end markers stay on the thread root exactly as Word writes them. The reply's `w15:commentEx/@w15:paraIdParent` points to the immediate parent's final-paragraph `w14:paraId`, so child-of-reply chains inherit the root range even though intermediate replies are reference-only. Both new entries begin reopened (`w15:done="0"`). |
| `UpdateComment(cmtAnchorId, markdown)` | Replace the body paragraphs; `w:id`/`w:author`/`w:initials`/`w:date` are untouched, and the old last paragraph's `w14:paraId` is re-stamped on the new last paragraph so Word's `commentsExtended` reply/resolve metadata stays attached across a body edit. |
| `SetCommentResolved(cmtAnchorId, resolved)` | Set Word's per-comment `w15:done` state (`true` resolves, `false` reopens). Calling it on a legacy flat comment upgrades that comment with the extension metadata without changing its body anchor or definition identity. |
| `RemoveComment(cmtAnchorId)` | Kind-guarded delegate to `DeleteBlock`'s `cmt` teardown: definition + marker triple (wrapper run included) everywhere in the package, plus threading pruning (below). |
| `ListComments()` | Read-only: `CommentListEntry(DefAnchorId, Author, Initials?, Date?, Text)` in comments-part order, with additive init-only `ParentAnchorId?` and `Resolved?` properties (the existing five-argument CLR constructor/deconstructor remains intact). `ParentAnchorId` maps `paraIdParent` back to the parent's stable `cmt` anchor. Both additive fields are absent/null when no `commentEx` entry exists, distinguishing a legacy flat comment from an explicitly reopened one. `Date` is the raw `w:date` string; `Text` is the flattened body (`annotationRef` runs excluded). The numeric `w:id` is never surfaced — comments are addressed by anchor. |

`Created` from `AddComment`, `AddCommentToRevision`, or `AddCommentReply` = the definition anchor (kind `cmt`) + its paragraph anchors
(kind `p`, scope `cmt`) — which are ordinary editable blocks, so `ReplaceText` on a comment
paragraph and `DeleteBlock` on the definition also work (pinned by `DS364`/`DS329`-style
tests). Body paragraphs only: Word has no comments-on-comments, so a `cmt`-scope host fails
with `AnchorWrongKind`; a zero-length span fails with the new `EmptyCommentSpan`.

### Threading metadata (`commentsExtended` / `commentsIds`)

Replies and resolve/reopen find or create the Word extension parts and stamp each participating
definition's final paragraph with `w14:paraId`. New para/durable ids are deterministic uppercase
eight-digit hex: max+1 across the ids already present in the package, with collision-free fallback
at the numeric ceiling. Existing entries, parentage, and done flags are preserved. A reply links
through `w15:paraIdParent`; `ListComments` resolves that para id back to the parent's `cmt` anchor.

These are relationship-bearing parts, not merely content blobs: `DocumentSnapshot` records their
relationship ids and XML, and restore reconciles part creation/deletion as well as content. Thus an
undo of the first reply/resolve removes parts that did not exist before, while redo recreates them
with their original relationship ids. Removing a comment still prunes its
`w15:commentEx`/`w16cid:commentId` entries and drops `w15:paraIdParent` attributes that pointed at
it (a surviving reply becomes top-level instead of dangling). That teardown remains in
`DeleteBlock`'s `cmt` branch so the generic delete path gets it too.

### `{#cmt:...}` payload tokens stay rejected

`CommentMarkerNotSupported` survives with a narrowed message naming `AddComment`: the
projection emits `{#cmt:...}` tokens only in the trailing `# Comments` section, never inline
in body text, so a token in an input payload has no round-trip meaning — the typed op is the
write path, exactly as `InsertFootnote` is for `[^label]`.

### Not yet

- **Cross-block spans.** A comment range is single-block in v1; Word can span paragraphs.
- **Anchor/span comments in header/footer/note stories.** Word allows them, but the
  `AddComment(anchorId, span, …)` form remains body-only, matching `InsertFootnote`'s host rule.
  `AddCommentToRevision` can target revision markup in every story `ListRevisions()` enumerates.
- **Tracked-changes mode.** Consistent with every other insert op, `AddComment` doesn't wrap
  its markup in `w:ins`.

## Tier E: Annotations

Anchor-addressed CRUD for the Docxodus annotation system (custom-XML +
bookmark pairs). Mutates the live session document; round-trips through
Save and Reopen. Exposed in .NET, WASM (`@docxodus/wasm`), and Python
(`docx-scalpel`).

### Methods

| Method | Description |
|--------|-------------|
| `AddAnnotation(anchorId, span?, DocumentAnnotation)` | Annotate a range; auto-generates 16-char hex id when `annotation.Id` is empty. `BookmarkName`, `AnnotatedText`, `Created`, and `PageInfoStale` are always set by the session. |
| `RemoveAnnotation(id)` | Removes the bookmark pair and the custom-XML entry. |
| `UpdateAnnotation(id, AnnotationUpdate)` | Mutates scalar fields (label/labelId/color/author) and metadata (per-key merge, explicit null = remove key). Range is preserved. |
| `MoveAnnotation(id, newAnchorId, newSpan?)` | Atomically re-targets to a new anchor + span. Validates the new range *before* removing the old bookmark. |

### `AnnotationUpdate`

```csharp
public sealed record AnnotationUpdate
{
    public string? LabelId { get; init; }
    public string? Label { get; init; }
    public string? Color { get; init; }
    public string? Author { get; init; }
    public IReadOnlyDictionary<string, string?>? MetadataPatch { get; init; }
}
```

Null/missing fields leave existing values unchanged. `MetadataPatch`
per-key semantics: non-null value = set/replace, explicit `null` = remove,
missing key = leave as-is.

### New error codes

- `DuplicateAnnotationId` — caller-supplied id already exists, or auto-id
  collided 4 times in a row (vanishingly rare).
- `AnnotationNotFound` — `Remove`/`Update`/`Move` invoked with an unknown id.
- `EmptyAnnotationSpan` — `span.Length == 0`, or `span == null` and the
  resolved block has zero inline runs.

### Return shape

All four ops use the standard `EditResult` envelope with one new field:
`AnnotationId` carries the affected id on success. `Created`/`Removed` are
always empty for these ops (the bookmark + custom-XML entry are internal,
not markdown anchors). `Patch` is null — annotation ops don't change the
markdown projection. `Modified` lists the enclosing block anchor (one
entry for Add/Remove/Update; one or two for Move depending on whether the
destination is the same block as the source).

## Find* helpers — anchor-level convenience over Grep

For anchor-level lookups that don't need match spans / fragments:

```csharp
session.FindByText(needle, options?)              // first anchor whose text contains needle, or null
session.FindAllByText(needle, options?)           // every anchor (deduplicated, in doc order)
session.FindByRegex(pattern, regexOptions?, opts?)// every anchor with at least one regex match
session.FindByKind("h", scope: "body")            // direct read over AnchorIndex, no text scan
```

`FindOptions { IgnoreCase, IgnoreWhitespace, KindFilter, ScopeFilter }`. `IgnoreWhitespace` flows down to `Grep`'s `WhitespaceMode.Normalize` so a needle written with regular spaces hits NBSP-using text — see the smoke-test trap that motivated #136 / #137.

## SmartQuotes

`DocxSessionSettings.SmartQuotes = true` makes every text-modifying op (`ReplaceText`, `ReplaceTextRange`, `ReplaceMatch`, `ReplaceTextAtSpan`) convert ASCII `"` and `'` in the payload to typographic curly quotes based on context — open at start/after-whitespace/after-open-bracket, close elsewhere. Avoids the cosmetic regression where a fill lands as `"foo"` next to surrounding already-curly `"foo"` text. Default off (pass payloads through unchanged).

## Switching tracked-changes mode mid-session (issue #304)

`Settings.TrackedChanges` and `Settings.RevisionAuthor` seed the session but do
**not** nail it down: `SetTrackedChanges(TrackedChangeMode)` and
`SetRevisionAuthor(string?)` switch how *subsequent* mutations are recorded, so a
workflow can edit directly and flip to `RenderInline` for just the edits that
should land as reviewable revisions (or change author between phases for
multi-author output) without a save→close→reopen.

Semantics:

- **Already-applied history is never touched.** Switching to `Accept` does not
  accept existing `w:ins`/`w:del` (that's `RevisionProcessor` / the MCP server's
  `accept_all`); switching to `RenderInline` does not retroactively wrap prior
  direct edits.
- **Session configuration, not a document mutation.** No undo snapshot is taken;
  `Undo()`/`Redo()` restore document content only and never change the mode or
  author. The mutators return `void`, not `EditResult`.
- Read back the current values via `session.TrackedChanges` / `session.RevisionAuthor`.
- Wired through every surface: WASM/npm (`setTrackedChanges`/`setRevisionAuthor`),
  stdio host + docx-scalpel (`set_tracked_changes`/`set_revision_author`), MCP
  (`docxodus_track_changes` action `set_mode`).

## Tracked-revision registry and resolution (issues #318, #455)

Tracked-change resolution used to be all-or-nothing (`RevisionProcessor` over the whole
document); the single most common review action — accept one revision, reject another —
required emulation. The session now exposes one live, part-aware registry and uses it for
individual and bulk resolution:

| Method | Description |
|--------|-------------|
| `ListRevisions()` | Read-only entries in document order across body, headers, footers, footnotes, and endnotes. Each carries an opaque stable `Id` (`rev2-…`), coarse `Type`, exact `Family`, native `ConstituentIds`, owning `PartUri`/canonical `Scope`, primary `AnchorId`, every `AffectedAnchor`, and a `ResolutionStatus` plus optional diagnostic. Authors/dates come from the live markup. Atomic entries include content and paragraph/row/property changes, named moves, cell insert/delete/merge operations, content-control envelopes, and numbering-property revisions. Unsupported, malformed, and ambiguous markup stays visible and fails closed. Legacy `revNNN` ids are accepted only as unambiguous inputs and are never emitted. |
| `AcceptRevision(id)` | Resolve ONE revision, keeping the change: unwrap `w:ins`/`w:moveTo`, carry out `w:del`/`w:moveFrom` (paragraph-mark deletions coalesce into the following paragraph, row deletions drop the row — the last row drops the table), drop the `*PrChange` element keeping current properties. An ordinary undoable mutation returning the `EditResult` envelope (`Modified` = touched blocks, `Removed` = blocks the resolution deleted). |
| `RejectRevision(id)` | The inverse: remove insertions, restore deletions (`w:delText` → `w:t`, marks stripped), keep a move at its source, restore a format change's stored old properties (preserving the children the `CT_*Base` inner schema excludes — mark revisions on a paragraph-mark `rPr`, header/footer references on `sectPr`, `rPr`/`sectPr` on `pPr`). |
| `AcceptAllRevisions()` / `RejectAllRevisions()` | Resolve the complete live registry through the same selective resolver as one atomic undo step. The registry is rebuilt after every entry so resolving a property shell can expose older archived revisions safely. Any unsupported, malformed, or ambiguous entry rolls back the whole operation. |

Mechanics live in `Docxodus/Internal/RevisionRegistry.cs` and `RevisionOps.cs`; the
per-element semantics mirror `RevisionProcessor`'s transforms, applied to one atomic group in
place. An unknown, already-resolved, or since-removed id fails with `RevisionNotFound`; unsafe
entries use `RevisionUnsupported`, `RevisionMalformed`, or `RevisionAmbiguous`. Wired through
every surface: WASM/npm, the stdio host + docx-scalpel, and MCP `docxodus_track_changes`.
The same ids can be passed to `AddCommentToRevision` (or `docxodus_comment add` with
`revisionId`) to anchor review discussion to the exact live change before it is resolved.

### Bulk resolution now fails closed — and that is a capability change

Before #455, `AcceptAllRevisions`/`RejectAllRevisions` were a whole-document
`RevisionProcessor` byte transform that always succeeded. They are now the selective resolver
run over every registry entry, and the resolver **refuses the whole operation** on the first
entry it cannot resolve safely. There is deliberately **no `force` mode**: a document whose
revision markup the registry does not understand cannot be bulk-resolved through any surface.

Concretely, these shapes were resolvable before and are refused now:

| Shape | Status | Diagnostic |
|-------|--------|-----------|
| A revision element with no `w:id` | `Malformed` | `missing_revision_id` |
| A revision element with a non-numeric `w:id` | `Malformed` | `invalid_revision_id` |
| One `w:id` shared by two distinct live groups in one part | `Ambiguous` | `duplicate_revision_id` |
| `w:customXmlMoveFromRange*`/`w:customXmlMoveToRange*` ranges | `Unsupported` | `unsupported_custom_xml_move_range` |
| `w:ins`/`w:del` inside an `m:ctrlPr` (math control properties) | `Unsupported` | `unsupported_revision_family` |
| `w:del` on a run's `w:rPr` or on a paragraph's `w:numPr` | `Unsupported` | `unsupported_revision_family` |
| A content-control (`w:sdt`) envelope whose range topology is not Word's two-pair shape | `Unsupported` | `unsupported_revision_family` |
| `w:numberingChange` not attached to `w:numPr` or a LISTNUM field | `Malformed` | `orphan_numbering_revision` |
| A cell marker that is not a direct `w:tcPr` property, or `w:cellMerge` without `w:vMerge` | `Malformed` | `orphan_cell_revision`, `invalid_cell_merge_state` |

`RevisionProcessor.AcceptRevisions`/`RejectRevisions` still handle all of them and remain
public, so a caller that needs the old always-succeeding behaviour can run the transform over
saved bytes and reopen the session. Whether the session surface should grow an explicit
opt-in escape hatch is an open public-API decision, not something the resolver should decide
silently.

## ApplyFormat — substring and TextMatch overloads

Three entry points for character-formatting (bold/italic/underline/strike/code/color/runStyle):

```
session.ApplyFormat(anchor, span, op)              // explicit CharSpan (use null for whole paragraph)
session.ApplyFormatToSubstring(anchor, str, op)    // find first occurrence of str, format it
session.ApplyFormat(textMatch, op)                 // exact span from a Grep result
```

The substring + `TextMatch` overloads exist because computing a `CharSpan` by hand is fragile when an auto-number prefix (`# Fourth The total…`) shifts the visible text relative to the run-text indices the `CharSpan` overload expects — see issue #138. Both convenience overloads just resolve to a `CharSpan` and call the underlying overload.

`ApplyFormatToSubstring` is named distinctly (rather than overloading) so existing `ApplyFormat(anchor, null, op)` whole-paragraph calls stay unambiguous to the C# resolver.

## FindPlaceholders — template-slot enumeration

`session.FindPlaceholders(kinds?, scope?)` is a thin classifier over `Grep` for the workflow every template-filling agent eventually writes itself. It scans for `\$?\[…\]` regions and tags each one as:

| `PlaceholderKind` | Pattern | What an agent does with it |
|---|---|---|
| `BlankFill` | `[___]` or `$[___]` (underscores only) | Fill with a literal value (a name, a number, a date) |
| `AlternativeClause` | `[clause text]` (anything else in brackets) | Keep, strip, or pick between alternatives |
| `Instruction` | `[insert X]`, `[specify Y]`, `[*italicized hint*]` | Parameter description — populate based on the hint |

`Instruction` placeholders expose the inner text (asterisks stripped) via the `Hint` field, so the agent can read `"insert percentage"` or `"specify name"` and decide what to put.

`PlaceholderKinds` is a flag enum (`BlankFill | AlternativeClause | Instruction = All`) for narrowing.

### The canonical fill recipe

```csharp
// Replace every value-blank in the document with a looked-up value.
foreach (var p in session.FindPlaceholders(PlaceholderKinds.BlankFill)
                          .OrderByDescending(p => p.Match.Span.Start))
{
    var value = LookupValueByContext(p.Match.ContextBefore, p.Match.ContextAfter);
    session.ReplaceMatch(p.Match, value);
}
```

This pattern — `FindPlaceholders` + `OrderByDescending(Span.Start)` + `ReplaceMatch` — collapses the 200-line context-needle-disambiguation script the Bluth-Co smoke test had to write down to a five-line loop. Process in reverse offset order so earlier-offset spans stay valid after later edits land, the same rule that applies to `ReplaceTextRange`'s internal pass.

### `FillPlaceholders` — picker-driven multi-pass fill

The 5-line recipe above is now a first-class call:

```csharp
var summary = session.FillPlaceholders(p => p.Kind switch
{
    PlaceholderKind.Instruction when p.Hint?.Contains("price") == true => "1.50",
    PlaceholderKind.BlankFill when p.Match.ContextBefore.TrimEnd().EndsWith("name is") => "ACME, INC.",
    _ => null
});
// summary.Filled / .Skipped / .StillPresent / .Passes / .Unfilled / .Errors
```

What `FillPlaceholders` does internally that the recipe doesn't:

- **Reverse-offset ordering.** Earlier-offset matches in the same paragraph go stale once a later edit lands; `FillPlaceholders` sorts every pass by `(anchorId desc, span.Start desc)` so each block's matches are applied right-to-left.
- **`$`-prefix preservation.** The placeholder regex `\$?\[…\]` captures `$[___]` including the leading `$`. With `FillOptions.PreserveDollarPrefix = true` (default), the picker's return value gets `$` prepended when needed so `"0.20"` lands as `$0.20`, not `0.20`.
- **Multi-pass iteration.** `FindPlaceholders` returns innermost brackets only; stripping the inner can surface a previously-nested outer. The loop re-finds placeholders each pass and stops when a pass makes zero changes (or `MaxPasses` — default 8 — is hit).

The picker is invoked for every kind in `FillOptions.Kinds`, which defaults to `PlaceholderKinds.All` — so a picker that wants to ignore alternative-clause brackets should return `null` for them rather than relying on the option to filter them out. Set `Kinds = BlankFill | Instruction` if you want the prior behavior of leaving alternative clauses untouched.

The picker is invoked once per placeholder per pass; return `null` to skip. `BulkEditResult.Unfilled` lists every placeholder the picker said `null` to (deduplicated across passes). `BulkEditResult.Passes` is the highest iteration pass that actually filled at least one placeholder (so a single-fill convergence reports `Passes = 1`, not 2).

`BulkEditResult.Skipped` is the *first-pass-null* count and is **not** a reliable "is the template done?" signal — a placeholder the picker said `null` to in pass 1 may be fully resolved by pass 2 (a nested-outer wrapper becomes fillable once its inner is stripped, or a structural delete removes the placeholder entirely). Assert on `BulkEditResult.StillPresent == 0` for the trustworthy single-call check: it's a post-loop `FindPlaceholders(opts.Kinds, opts.Scope).Count`, so `Skipped > 0 && StillPresent == 0` correctly reads as "picker skipped on the first pass but later passes finished the job."

#### `FillOptions.CoalesceWhitespaceAroundEmptyFill`

Returning `""` from the picker is the canonical "drop this optional clause entirely" signal. By default the bracketed match is deleted exactly, which leaves whitespace and punctuation around the (now-gone) brackets untouched. The repro from issue #188:

```
... pursuant to the General Corporation Law on March 14, 2024 [under the name [_______________]].
```

For a company that has never been renamed, the picker fills the inner name slot with `"Bluth, Inc."`, then in pass 2 sees the outer wrapper `[under the name Bluth, Inc.]` and returns `""`. The literal-delete result is:

```
... pursuant to the General Corporation Law on March 14, 2024 .
```

Note the stray space before the period.

With `FillOptions.CoalesceWhitespaceAroundEmptyFill = true`, an empty fill (after `$`-prefix preservation has been applied) absorbs adjacent chars based on the immediate neighbors of the placeholder span:

| Left neighbor | Right neighbor | Action | Example |
|---|---|---|---|
| whitespace (incl. NBSP, narrow NBSP, thin space) | whitespace | consume the trailing space, leaving one | `"alpha [x] beta"` → `"alpha beta"` |
| whitespace | `.` `,` `;` `:` `!` `?` | drop the leading space | `"… 2024 [x]."` → `"… 2024."` |
| `(` `[` `{` | matching `)` `]` `}` | drop both surrounding brackets | `"[[x]]"` → `""` |

Default is `false` (preserve the literal-delete behavior). `$`-prefix preservation runs first, so a picker returning `""` for `$[xxx]` with the default `PreserveDollarPrefix = true` ends up replacing with `"$"` (not empty) and coalescing is skipped — set `PreserveDollarPrefix = false` when you want the `$` to drop along with the brackets.

Note the .NET implementation reads the live flat text of the enclosing block to find the immediate neighbors, so the rules work regardless of the `Boundary` setting. The npm TypeScript implementation uses the match's already-populated `contextBefore` / `contextAfter` (zero extra round-trip) — with `boundary: ContextBoundary.Bracket` the outer brackets are not visible to context, so the bracket-coalesce rule won't fire on the JS side. Callers using `CoalesceWhitespaceAroundEmptyFill` should leave `Boundary` at the default `Char`.

### `ReplaceInner` — strip brackets while preserving prefix/suffix

```csharp
// match.Text = "$[___]"
session.ReplaceInner(match, "0.20");
// paragraph now contains "$0.20" (the leading $ outside the brackets survives).
```

`ReplaceInner` parses the brackets out of `match.Text` and substitutes the new inner for everything between (and including) `[` and `]`, then dispatches to `ReplaceMatch` with the recomposed string. Returns `MalformedMarkdown` if the match text has no balanced brackets. The shared `DocxSessionOps.ReplaceInner` is reused by both the WASM bridge and the stdio NDJSON host, so the `replace_inner` op is available to Python wrappers too.

### `TemplatePlaceholder.AlternativeKinds`

When the primary classification is borderline, secondary classifications are exposed via the `AlternativeKinds` list. The current borderline case: a `BlankFill` whose inner text contains 4+ spaces (i.e. reads like a multi-word clause that happens to contain a `_______`). Primary `Kind` stays `BlankFill` for back-compat; `AlternativeKinds` lists `AlternativeClause` so callers can detect the ambiguity.

### Nesting

Nested brackets (e.g. `[under the name [Bluth, Inc.]]`) resolve to the INNERMOST bracket only — usually what the agent cares about, since the inner is the value slot and the outer is the optional-clause wrapper. If you need both, use `Grep` directly with a balanced-bracket pattern.

## Edit-state introspection — `GetEditSummary` and `GetDiff`

### `GetEditSummary` — "am I done?"

`session.GetEditSummary()` returns a single `EditSummary` record composing
existing primitives:

| Field | Source |
|---|---|
| `TotalAnchors` | `Project().AnchorIndex.Count` |
| `RemainingPlaceholders` | `FindPlaceholders(All, All)` |
| `BareUnderscoreRuns` | `Grep(@"(?<![\[_])_{3,}(?![\]_])")` (underscore-aware lookarounds bound the maximal run so the count matches the visible underline groups — see DS280b/c) |
| `FootnoteCount` | `AnchorIndex` filter on `kind=fn, scope=fn` (excludes reserved separators per #162) |
| `InlineFootnoteRefCount` | Body part's `w:footnoteReference` count |
| `CommentCount` | `AnchorIndex` filter on `kind=cmt` |

Designed to make verification logic declarative:

```csharp
var summary = session.GetEditSummary();
Assert.Empty(summary.RemainingPlaceholders);
Assert.Empty(summary.BareUnderscoreRuns);
Assert.Equal(0, summary.FootnoteCount);  // commentary stripped
```

### `GetDiff` — "show me what I changed"

`session.GetDiff(DiffFormat.Json)` (default) returns an anchor-keyed JSON
array of `DiffEntry` records comparing the projection captured at session
construction time against the current state.

```json
[
  { "op": "delete", "anchorId": "p:body:abc…", "before": "Drafting Note..." },
  { "op": "modify", "anchorId": "p:body:def…", "before": "[___]", "after": "ACME, INC." },
  { "op": "insert", "anchorId": "p:body:ghi…", "after": "New paragraph text" }
]
```

Initial-projection capture is on by default (`DocxSessionSettings.CaptureInitialProjection = true`)
and costs ~200ms at construction. Turn it off if you don't plan to diff.

`DiffFormat.Unified` returns a `patch(1)`-compatible unified diff over the
markdown projections (`--- initial` / `+++ current` headers, 3 lines of
context per hunk, hand-rolled LCS over `\n`-split lines). Returns the empty
string when nothing has changed:

```diff
--- initial
+++ current
@@ -1,6 +1,6 @@
 # Document

-{#p:body:6b6439…} First paragraph.
+{#p:body:6b6439…} REPLACED PARAGRAPH

 {#p:body:a321f0…} Second paragraph.
```

`DiffFormat.SideBySide` returns a `diff -y`-style two-column rendering — the
initial projection padded to 72 chars on the left, a one-character marker
(`' '` unchanged, `'|'` modified, `'<'` only-initial, `'>'` only-current),
then the current projection on the right. Adjacent `Delete + Insert` pairs
collapse to a single `|` "modified" row.

Both line-based formats diff the raw markdown projection — anchor tokens
(`{#…}`) appear verbatim in the output. Switch to
`AnchorIdRendering = Abbreviated` or `Sequential` in
`DocxSessionSettings.ProjectionSettings` to keep token noise out of the
diff when reviewing in a terminal.

## Sliced projection — `ProjectAnchor`

`session.Project()` returns the full document — usually overkill when an agent
only needs to read or edit one section. `session.ProjectAnchor(anchorId, depth?)`
returns a `MarkdownProjection` whose `Markdown` contains only the blocks in
scope and whose `AnchorIndex` is filtered to those blocks plus their
descendants:

```csharp
// Just the heading paragraph itself
var self = session.ProjectAnchor(headingAnchor, ProjectionDepth.SelfOnly);

// A table and all its rows/cells
var table = session.ProjectAnchor(tblAnchor, ProjectionDepth.Subtree);

// A heading + everything under it up to the next same-or-higher heading
// (the default — the "give me this section" case)
var section = session.ProjectAnchor(headingAnchor);
```

`ProjectionDepth` values:

| Value | Behavior |
|---|---|
| `SelfOnly` | Just the addressed block — its anchor and its own text. For headings, returns only the heading paragraph, not the section underneath. |
| `Subtree` | Self + descendants. Most useful for `tbl` anchors (returns the whole table). For paragraph-like anchors, equivalent to `SelfOnly` since they have no descendants. |
| `SubtreeAndFollowingSiblings` (default) | Self + descendants + following siblings up to (but not including) the next sibling at the same or higher heading level. For non-heading anchors, equivalent to `Subtree`. |

Useful for showing an LLM one section at a time without paying the ~1 s
full-projection cost per turn — the agent reads, decides, edits, and
re-projects only the slice it touched.

The `anchorId` argument accepts whatever rendering form the projection's
`AnchorIdRendering` setting emits — the dual-keyed `AnchorIndex` resolves
full Unids, abbreviated ids, and sequential ids interchangeably. See
[`markdown_projection.md`](markdown_projection.md#anchor-id-rendering-modes)
for the rendering modes.

Returns the same `MarkdownProjection` shape as `Project()` — caller code that
already consumes the full projection (e.g., reading `AnchorIndex` to find
follow-up edit targets) works unchanged on a slice. Throws
`InvalidOperationException` if the anchor isn't in the current `AnchorIndex`.

## ReplaceTextRange — surgical text edits

`session.ReplaceTextRange(anchorId, find, replace, options?)` finds every literal occurrence of `find` in one paragraph/heading/list-item's flat text and substitutes `replace` for each, returning an `EditResult` per attempted match. Built on `Grep` — same fragment walker, opposite direction.

Three entry points covering the natural workflows:

```
session.ReplaceTextRange(anchor, find, replace, opts?)         // most common: replace every match in one block
session.ReplaceMatch(textMatch, replace)                       // convenience for a Grep result
session.ReplaceTextAtSpan(anchor, spanStart, spanLength, repl) // exact-span variant when several identical needles share a block
```

`ReplaceOptions`: `IgnoreCase` (case-insensitive find), `MaxReplacements` (cap on
how many to apply), `ExpectedMatchCount` (require the exact live count before the
cap), and `Preconditions` (the common optimistic guard object).

### Formatting-preservation contract

The replacement text inherits the formatting of the FIRST run the match spanned. Middle and trailing runs keep their `w:rPr` but lose the slice of text the match consumed — so a bold phrase that got partially overwritten still has bold formatting for any surviving text. This is the practical sweet spot: it solves the template-fill case where you want `[___]` → `Bluth Co.` to take on the surrounding text's formatting, and it's predictable for cross-formatting matches.

If you need different per-fragment behavior (e.g., the replacement should be bold even when the first fragment was plain), use `Grep` + bespoke `Raw.GetXml` mutation today, or wait for a future inline-markdown-aware overload.

### Tracked-change behavior

With `TrackedChanges = RenderInline`, all three entry points retain the same surgical boundaries but record them as native Word revisions. Text before and after the span remains in ordinary runs; the selected slices become formatting-preserving `w:r/w:delText` children under `w:del`; and the replacement becomes one `w:r/w:t` under `w:ins`, carrying the first affected run's `w:rPr`. The envelopes use the session revision author, operation timestamp, and fresh revision ids.

Revision wrappers stay inside a match's hyperlink, run-level SDT, `smartTag`, or `fldSimple` container. A match that crosses formatting boundaries keeps one deleted run per source format. Zero-width bookmark/comment/permission/proofing markers and footnote/endnote reference runs remain outside the revision envelopes, so accepting or rejecting the edit cannot silently destroy their relationships. `AcceptRevision`/`RejectRevision`, whole-document accept/reject, and undo/redo therefore resolve surgical edits the same way as full-block tracked replacements. `TrackedChanges = Accept` retains the original direct run-mutation behavior.

### Ordering and atomicity

Multiple matches in the same paragraph are applied in **reverse document order** so each earlier-offset match's span stays valid after later edits land — the same trick the projector uses for tracked-change accept passes. The whole call records **one** snapshot; `Undo()` rolls every replacement back together. Preconditions, exact occurrence counting, and the rewrite execute under one mutation gate; a count mismatch returns one failed result and leaves bytes, version, and undo history unchanged.

### When to reach for the span-addressed variant

If the agent has computed five `[___]` placeholder matches in the same paragraph from `Grep` and wants to fill each with a different value, `ReplaceTextRange` would only see "five identical `[___]` needles" and replace each with the first value (or all with the same value). `ReplaceTextAtSpan` (or `ReplaceMatch`) addresses each match by its exact coordinates so the disambiguation is unambiguous. Apply spans in **reverse offset order** in this case for the same reason — earlier spans stay valid after later edits.

### Recipe: enumerate-and-fill via Grep + ReplaceMatch

```csharp
foreach (var match in session.Grep(@"\[_+\]")
                             .OrderByDescending(m => m.Span.Start))
{
    var value = LookupValueByContext(match.ContextBefore, match.ContextAfter);
    session.ReplaceMatch(match, value);
}
```

This pattern collapses the 200-line context-needle-disambiguation script the Bluth-Co smoke test had to write down to a five-line loop.

## Grep — cross-run text search

`session.Grep(pattern, options?, scope?, contextChars?)` searches the flat text of every paragraph/heading/list-item in scope, returning matches in document order. Each `TextMatch` carries:

- `EnclosingAnchor` — the smallest block-level anchor that fully contains the match.
- `Span` — character offset+length within the enclosing block's flat text.
- `Fragments` — one `RunFragment` per `<w:r>` the match spans, in document order. Each fragment names the run's Unid, the slice of the match it contributes, the offset+length inside the run, and the run's visible `Formatting` (bold/italic/strike/underline/code/color/hyperlink/runStyle).
- `ContextBefore` / `ContextAfter` — up to `contextChars` (default 40) of surrounding text from the same block.
- `Groups` — regex capture groups.

The fragment breakdown is the whole point: Word splits paragraph text into many `<w:r>` elements at every formatting boundary, so a placeholder like `[_______________]` routinely spans 2–3 runs. Without the fragment list, an agent doing search/replace has to either flatten runs (losing per-fragment formatting) or skip split matches (missing real text). `Grep` does the walk once and hands back the breakdown so callers can preserve each fragment's formatting when rewriting.

### When to use

```
Need to … → use
Find every literal/regex pattern in the doc → Grep
Find one anchor whose text contains X → Grep, take .First().EnclosingAnchor
Enumerate template placeholders → Grep(@"\[_+\]") or similar
Edit text without losing formatting → Grep + a fragment-aware rewrite (see #139 for the planned ReplaceTextRange built on this)
Find a multi-paragraph clause or pattern that straddles a paragraph break → GrepCrossBlock
```

### Performance

~400 ms for a full-document grep over the 150 KB NVCA Model COI (~500 anchors, ~31 underscore-placeholder matches). Scales linearly with document size + match count.

### Known limits

- **Each block is grep'd in isolation.** Grep iterates paragraphs/headings/list-items and runs the regex against each one's flat text independently. `session.Grep("Hello world")` won't match if `"Hello "` is in one paragraph and `"world"` is in the next, even though they appear adjacent in the rendered doc. This is by design: every `TextMatch` carries a single `EnclosingAnchor` for the caller to hand back to `ReplaceText`/`Raw.ReplaceXml`. For cross-block search (legal clauses split for readability, multi-paragraph regions, etc.) use **`GrepCrossBlock`** (see next section).
- `RegexOptions` is the .NET enum; the npm wrapper passes its numeric value through (see `GrepOptions` in `npm/src/types.ts`).
- Tracked-change content currently follows the projector's accepted/rendered text — `Settings.TrackedChanges = StripDeletions` won't filter `<w:del>` content out of Grep yet. Worth opening as a follow-up if it matters.

### Context boundary modes

`Grep` / `GrepCrossBlock` / `FindPlaceholders` accept a `ContextBoundary` parameter
that decides where the context-computation walker stops:

| Mode | Stops at | Use when |
|---|---|---|
| `Char` (default) | nothing — truncate at `contextChars` | legacy callers, free-form text where boundaries are noisy |
| `Bracket` | `[`, `]` | template fills with adjacent placeholders — each `ContextBefore`/`ContextAfter` is guaranteed to belong to this match only |
| `Sentence` | `.`, `!`, `?`, `:`, `;` | LLM prompt-building where each snippet should be a self-contained sentence |
| `Comma` | `,` | matches inside enumerations |

The default `contextChars` widened from 40 → 80 in #164. Combined with `Bracket`
mode this lets a template-fill picker use plain `.Contains` / `EndsWith` checks
without cross-pollution from adjacent placeholders:

```csharp
var matches = session.Grep(@"\[CITY\]",
    scope: ProjectionScopes.Body,
    contextChars: 80,
    boundary: ContextBoundary.Bracket);
// matches[0].ContextBefore guaranteed bracket-free
```

## GrepCrossBlock — cross-block text search

`session.GrepCrossBlock(pattern, options?, scope?, contextChars?, whitespace?)` is the variant of [`Grep`](#grep--cross-run-text-search) for matches that legitimately span multiple paragraphs — legal clauses split across paragraphs for readability, multi-paragraph indemnification blocks, or `Section \d+\.\d+\b` straddling a paragraph break.

Each `CrossBlockMatch` carries:

- `Text` — the matched text, with single `\n` characters at each block boundary the match crossed.
- `EnclosingAnchors` — every block-level anchor the match touches, in document order. Always non-empty.
- `Slices` — per-block breakdown. Each `BlockSlice` names its `Anchor`, the `SpanInBlock` (offset+length within that block's own flat text), and a `Fragments` list with the same shape as `Grep`'s.
- `ContextBefore` / `ContextAfter` — surrounding text from the concatenated stream; may include block-boundary `\n` characters.
- `Groups` — regex capture groups.

### Separator and regex behavior

Adjacent blocks in the searched text are joined with a single `\n`. That means:

- `^` and `$` with `RegexOptions.Multiline` anchor at block boundaries.
- `.` does not match across boundaries unless `RegexOptions.Singleline` is set.
- `\s`, `\n`, and explicit `\n` patterns in your regex see the boundary.

### What it never crosses

Matches are scoped strictly to keep them meaningful for downstream editing:

- **Package parts** — body → footnote, header → body, etc. Different package parts are searched independently.
- **Container boundaries** — a body paragraph cannot bridge into a table-cell paragraph. Table cells form their own groups (`w:tc` is the parent).
- **Non-paragraph siblings** — a `w:tbl`, `sectPr`, or any non-`w:p` element between two paragraphs breaks the run; matches don't bridge across it.

### Superset of `Grep`

A single-block match still appears in the results with one `Slice`. Filter `Slices.Count > 1` if you only want cross-block hits. The naming reflects "the variant that also handles cross-block," not "only cross-block."

### Edit semantics — deferred

Replace on a cross-block match has at least three reasonable behaviors (merge into one block, per-slice independent rewrites, boundary-preserve), none obviously right. Edit primitives are deliberately out of scope until a concrete consumer surfaces the right semantics. Today, callers can read the slice list and apply slice-by-slice edits via `ReplaceTextAtSpan` themselves.

### Performance

Same order of magnitude as `Grep`: one concatenation pass + one regex pass per sibling group, with `RunTextMap` shared for fragment resolution. Memory grows with the largest group's concatenated text, not the whole document.

## The Raw escape hatch

`session.Raw` exposes three operations: `GetXml(anchorId)` returns the element's OOXML as a string (useful as a template), `InsertXml(anchor, position, xml)` inserts a sibling fragment, `ReplaceXml(anchor, xml)` swaps the element for a fragment. Newly-inserted elements automatically get Unids and become addressable on the next projection.

The validation pipeline is short-circuit ordered: well-formedness (`MalformedXml`), namespace whitelist check (`DisallowedNamespace` — only `w:`, `m:` for math, `wp:`/`a:` for drawing, `r:`, and our own PtOpenXml namespace are allowed), structural slot check (`IncompatibleElementType`). Setting `Settings.ValidateRawOps = true` additionally runs `OpenXmlValidator` before and after the op and rolls back via the snapshot if the post-op error count is greater than the pre-op count. Pre-existing schema issues in the input document are tolerated (the validator is only used to detect deltas, not to gate the document overall), and the projector's internal `PtOpenXml:Unid` attributes are filtered out before counting since they are not in the OOXML schema by design. Slower than the default path but bulletproof for untrusted agent input.

**The round-trip recipe.** This is the safe pattern for raw mutations the agent should always prefer over authoring XML from scratch:

```csharp
// .NET
var xml = session.Raw.GetXml(anchor);
var modified = MutateSomehow(xml);
var result = session.Raw.ReplaceXml(anchor, modified);
```

```typescript
// TypeScript
const xml = session.raw.getXml(anchor);
const modified = mutateSomehow(xml);
const result = session.raw.replaceXml(anchor, modified);
```

Starting from a known-valid XML fragment and modifying it locally is dramatically less error-prone than constructing OOXML from scratch — namespace declarations, attribute ordering, and child-element validity are all preserved from the original.

## Canonical table addressing

The canonical address for every cell operation is the physical `w:tc` anchor (`tc:{scope}:{unid}`).
`InsertTable` and cell-creating mutations return `tc` anchors in `Created`; callers do not need to
search for a paragraph merely to identify its cell. `GetTableMetadata(tblAnchor)` returns the table
anchor and explicit ordered row, column, and physical-cell metadata. A cell records its row index,
starting grid column, horizontal/vertical spans, vertical-merge role, owning table/row identities,
and only its **direct** paragraph anchors. Paragraphs in a nested table belong solely to that nested
table's cells.

Resolution is deliberately bidirectional:

- `ResolveTableCellAnchor(tcAnchor)` returns the cell's current coordinate and spans.
- `ResolveTableCellCoordinate(tblAnchor, rowIndex, columnIndex)` returns the physical cell covering
  that Word-grid coordinate, including a horizontally spanned cell. A coordinate in a
  `gridBefore`/`gridAfter` gap returns `AnchorNotFound`; it never guesses a neighboring cell.

Compatibility is narrow and deterministic. A legacy `p`/`h`/`li` anchor whose nearest ancestor is
a cell is translated to that nearest `tc`, so old callers have a migration window and nested tables
cannot retarget an outer cell. Passing `tbl`, `tr`, or an unrelated paragraph to a cell operation
returns `TableAnchorMigrationRequired` with instructions to call `GetTableMetadata` or coordinate
resolution. New code should never cache or manufacture cell-paragraph addressing.

Every table-shape mutation populates `EditResult.TableAnchors`:

- `Retained` pairs each stable identity's before/after grid location;
- `Added` lists new table/row/column/cell identities at their new locations;
- `Invalidated` lists identities that no longer resolve at their former locations.

The lists are deterministic (old-location order for retained/invalidated, new-location order for
added), so clients can update a cached coordinate model without matching by array position.
`Created`/`Removed` remain the concise mutation result and use canonical `tc` identities;
`TableAnchors` is the complete structural account.

Real columns are the `col` anchors of `w:tblGrid/w:gridCol` and retain their Unids when widths or
neighboring columns change. A table with a missing or underspecified `tblGrid` is not mutated by a
read: metadata derives deterministic virtual `col` identities (`IsVirtual = true`). The first
column/width transaction materializes real `gridCol` elements inside that transaction and reports
the virtual columns invalidated and the real columns added. Persist identities across a close/reopen
checkpoint with `Save(persistAnchorIds: true)` (or the equivalent session setting); a normal clean
save intentionally strips all Unid bookkeeping, including table identities.

`DocumentStructure` retains its path-based `Id` as a compatibility/display locator and adds
`AnchorId` for addressable table/row/cell elements. `TableColumnInfo` similarly carries both legacy
path ids and canonical column/table/cell anchors plus `IsVirtual`. Its coordinates use this same
grid model, so `gridBefore`/`gridAfter`, `gridSpan`, and actual `vMerge` runs cannot diverge from the
live session APIs.

## Table cell merge: the grid model

`MergeCells`/`UnmergeCells` (issue #340 Stage B) and the row/column CRUD around them share one
model, and every rule below falls out of it.

**Geometry.** A row's cells tile `w:tblGrid` columns left→right. Each cell covers `w:gridSpan`
columns (default 1), starting from an origin shifted by `w:trPr/w:gridBefore`. A vertical merge is
a *column-aligned run of rows* whose lead cell carries `w:vMerge w:val="restart"` and whose
followers carry a bare `w:vMerge`. So a cell has a grid rectangle, and "which cell is in column
N" is a lookup over that geometry — never an index into `w:tr`'s children. That distinction is the
whole fix: the pre-#340 CRUD indexed cells positionally, which silently tore the grid the moment a
span existed.

**`MergeCells(cellAnchor, rowSpan, colSpan, options?)`.** The rectangle starts at the anchor's cell
and runs `rowSpan` rows down × `colSpan` *cells* right (measured in the anchor row, so it composes
with spans already there). It is applied only if:

- it stays inside the table (`rowSpan`/`colSpan` in range) and covers ≥ 2 cells;
- every covered row tiles *exactly* the same grid columns `[c0, c1)` — otherwise an existing
  `w:gridSpan` straddles the rectangle's edge and merging would leave the grid ragged;
- its first row is not itself a `w:vMerge` continuation, and no continuation follows its last row.

Each violation is `InvalidTableMerge` with a message naming which one. Nothing is half-applied —
validation runs before the undo snapshot is taken.

The result: each covered row keeps its first cell, drops the rest, and gets
`w:gridSpan = c1 - c0` plus a `w:tcW` summed over the grid columns it now covers. With `rowSpan > 1`
the lead cell gets `w:vMerge w:val="restart"` and the rows below a bare `w:vMerge`.

**`UnmergeCells(cellAnchor)`** is the inverse: drop `w:gridSpan`/`w:vMerge`, restore one cell per
grid column (cloning the merged cell's shell — borders, shading, valign — minus the merge markup)
and give each its `w:tblGrid` width. Addressing a *continuation* cell unmerges the whole run: the
op walks up to the restart and back down through every column-aligned continuation. A cell with no
merge markup is `InvalidTableMerge`, not a silent no-op.

**Anchor semantics.** Content blocks keep their own identities when moved, while physical cell
shells follow the canonical structural lifecycle:

| Cell | What happens |
|---|---|
| The surviving (lead) cell | Its `tc` identity is retained and returned in `Modified` |
| An absorbed cell, `Content = Append` (default) | Non-empty blocks are **moved** into the lead cell with their Unids; the absorbed `tc` identity is invalidated and returned in `Removed`/`TableAnchors.Invalidated` |
| An absorbed cell, `Content = Discard` | Its content is dropped and its `tc` identity is invalidated |
| An absorbed cell, `Content = Reject` | Nothing happens — a non-empty absorbed cell fails the whole op |
| A vertical-merge continuation | The `tc` survives at the same coordinate and is retained; its body is reduced to the one empty `w:p` CT_Tc requires |

A continuation `tc` stays addressable even though Word renders its body invisibly; unmerge before
writing content intended to display. A horizontal append merge invalidates absorbed cell shells even
though their content blocks survive inside the lead cell.

**Projection.** A table carrying any merge fails the projector's GFM-simplicity predicate (any
`w:gridSpan > 1` or any `w:vMerge` disqualifies), so it renders as the opaque ` ```table ` block
with its `{#tbl:…}` anchor — and every surviving cell stays individually addressable in the anchor
index. Use `ReplaceCellContent(tc, …)` for whole-cell content, or a direct paragraph anchor from
table metadata for paragraph-grained `ReplaceText`/`ApplyFormat`.

**Span-aware CRUD.** The four reshaping ops each have one defined behavior where a merge is in the
way — extend, narrow, or repair, never tear:

| Op | Against a merge |
|---|---|
| `InsertTableRow` | Mirrors the reference row's grid shape (widths + `w:gridSpan`), never its merge markup. Where a vertical merge *crosses* the insertion boundary — the row on the far side of it carries a continuation — the new row joins the run as a continuation, so the merge extends instead of being punched through |
| `DeleteTableRow` | Deleting a merge's lead row promotes the next row's continuation to the new `w:vMerge w:val="restart"`, so the run is never left headless |
| `InsertTableColumn` | A cell straddling the new boundary widens by one column (`gridSpan + 1`, width grown by the new `w:gridCol`) instead of gaining a sibling; rows whose cells end at the boundary get an ordinary new cell |
| `DeleteTableColumn` | A cell spanning the doomed column narrows by one (`gridSpan - 1`, width shrunk by the column) instead of disappearing; unit cells covering it are removed as before |

`SetColumnWidths` follows the same rule: a merged cell's `w:tcW` is the **sum** of the grid columns
it spans, not the width at its position in the row.

**Downstream.** `IrReader` already models `gridSpan`/`vMerge` (`IrCell.GridSpan`/`IrVMerge`, folded
into `IrCell.ShellDigest`), so merged documents flow through `DocxDiff` with the usual round-trip
contract — accept ≡ right, reject ≡ left (test `DT240`).

## Error catalog (by remediation)

Errors are grouped by what the agent should do in response, not by where in the code they're raised. The `EditErrorCode` enum lives in `Docxodus/DocxSession.cs`; the snake-case TypeScript union is in `npm/src/types.ts`.

| The agent should… | When it sees these codes |
|---|---|
| Re-read the current version/target metadata in `error.precondition`, rebase or abandon the stale edit, then retry with fresh guards | `PreconditionFailed` |
| Re-project and re-derive the anchor from current text | `AnchorNotFound` |
| Re-list revisions (`ListRevisions`) and reissue with a current id | `RevisionNotFound` |
| Re-list native objects and reissue with a current id/name | `HyperlinkNotFound`, `BookmarkNotFound` |
| Inspect the revision diagnostic and repair/reopen the source document | `RevisionUnsupported`, `RevisionMalformed`, `RevisionAmbiguous` |
| Re-read the anchor's kind via `GetAnchorInfo`, reissue with the right op or coordinates | `AnchorWrongKind`, `TableAnchorMigrationRequired`, `AnchorsNotAdjacent`, `InvalidPosition`, `OffsetOutOfRange`, `EmptyCommentSpan`, `EmptyHyperlinkSpan` |
| Fix the target/name or resolve the existing reference first | `DuplicateBookmarkName`, `InvalidBookmarkName`, `InvalidHyperlinkTarget`, `MissingBookmarkTarget`, `BookmarkInUse`, `ManagedBookmark` |
| Choose a safe run/range boundary, resolve the pending structural revision, or switch subsequent edits out of tracked mode | `UnsupportedInlineBoundary`, `UnresolvedStructuralRevision`, `TrackedOperationUnsupported` |
| Fix the markdown payload (the message names what's wrong) | `MalformedMarkdown`, `UnsupportedMarkdownSyntax`, `AnchorTokenInPayload` |
| Call the dedicated op the message names (`InsertTable`, `InsertFootnote`/`InsertEndnote`, `AddComment`, `InsertImage`), or fall back to `Raw.InsertXml` | `TableInsertNotSupported`, `FootnoteRefNotSupported`, `CommentMarkerNotSupported`, `ImageInsertNotSupported` |
| Re-query `ListStyles()` for a current style id, or `GetListMembership()` for the valid numbering level | `UnknownStyle`, `InvalidListLevel` |
| Fix the op's field values (the message names the constraint OOXML can't express) | `InvalidPageNumbering`, `InvalidParagraphFormat`, `InvalidListStartValue`, `InvalidTableStyling`, `InvalidTableMerge` |
| Use `Raw.GetXml(anchor)` as a template, mutate, resubmit | `MalformedXml`, `DisallowedNamespace`, `IncompatibleElementType`, `ValidationFailed` |
| Stop, reopen, or accept "no more history" | `SessionDisposed`, `NothingToUndo`, `NothingToRedo` |
| Should not happen; treat as a bug. Op is rolled back, safe to retry once or report. Full exception is on `session.LastInternalError` | `InternalError` |

For batched lookups (an agent that just enumerated 50 anchors and wants
previews for all of them), use `session.GetAnchorInfos(ids)` — a single pass
over the AnchorIndex instead of one walk per id. Returns
`IReadOnlyDictionary<string, AnchorInfo?>` — unknown ids map to null.

**Failure is transactional.** On any error, no mutation was applied. The pre-op snapshot was taken but is discarded without restoring (because nothing landed in the first place). Failed ops do not consume an undo slot. This holds for both pre-apply validation failures and runtime failures caught and rolled back.

## Recipes

These are worked examples drawn from the end-to-end smoke test (`DocxSessionSmokeTest.cs::DS999`) and the per-tier tests, lightly genericized. They use the .NET API; the TypeScript API is shape-identical (camelCase method names, `string` anchors, `Promise`-free synchronous returns from the npm wrapper since everything runs on the WASM worker).

### Inspect before editing

Do not guess style ids, re-create inherited formatting as direct XML, or infer run boundaries from
markdown. Query the live document, then feed the returned identifiers and coordinates back into the
matching mutation API unchanged:

```csharp
var style = session.ListStyles().Single(s => s.Name == "Strong Custom");
var formatting = session.GetFormatting(paragraphAnchor)!;
var word = session.ListInlineSpans(paragraphAnchor).Single(s => s.Text == "Defined Term");

// style.Id is the document's real w:styleId; word.AnchorId + word.Span is ApplyFormat-ready.
session.ApplyFormat(word.AnchorId, word.Span, new FormatOp { RunStyle = style.Id });
```

`DirectParagraph`/`InlineSpan.Direct` say what is written on the target itself. Their
`EffectiveParagraph`/`Effective` counterparts say what Word renders after document defaults and
the complete style chain are applied by `FormattingAssembler`. Keeping those two layers separate
is essential: an absent direct value means “inherit,” not false or zero.

### Replace a clause's text while preserving its style and numbering

```csharp
using var session = new DocxSession(docxBytes);
var anchor = session.Project()
    .AnchorIndex.Values
    .First(t => t.Anchor.Kind == "h" && t.Anchor.Scope == "body")
    .Anchor.Id;

var result = session.ReplaceText(
    anchor,
    "**Indemnification.** The Provider shall indemnify the Client for any [breach](https://example.com/terms#breach) of the foregoing.");

// result.Success == true
// result.Modified[0].Id == anchor   (kind/scope unchanged)
// result.Patch.Markdown contains the freshly-projected scope
// The paragraph's existing w:pPr (Heading1 style + numbering)
// is preserved — only the runs were swapped.
```

### Split a paragraph and promote the second half to a heading

```csharp
var split = session.SplitParagraph(originalAnchor, characterOffset: 42);
// split.Modified[0].Id == originalAnchor   (first half keeps the Unid)
// split.Created[0]    is the new anchor on the second half

var secondHalf = split.Created[0].Id;
session.SetParagraphStyle(secondHalf, "Heading2");
// The anchor's kind prefix is now 'h' instead of 'p';
// resolution by Unid still works either way.
```

### Format a character range with bold

```csharp
// Bold characters 0..5 of the paragraph (whole-paragraph: pass null span)
var r = session.ApplyFormat(
    anchor,
    new CharSpan(0, 5),
    new FormatOp { Bold = true });
```

### Inject a content control via raw XML

```csharp
var xml = session.Raw.GetXml(paragraphAnchor);
// Wrap the paragraph in a w:sdt for structured tagging
var modified = WrapInSdt(xml, tag: "PartyName", alias: "Party Name");
var r = session.Raw.ReplaceXml(paragraphAnchor, modified);
// r.Created includes the SDT and the preserved inner paragraph anchors
```

### Apply edits as tracked revisions instead of accepted changes

```csharp
var settings = new DocxSessionSettings
{
    TrackedChanges = TrackedChangeMode.RenderInline,
    RevisionAuthor = "agent-alpha",
};
using var session = new DocxSession(docxBytes, settings);

session.ReplaceText(anchor, "Updated clause text.");
// The document now contains <w:del> wrapping the old runs and
// <w:ins> wrapping the new runs. The anchor stays live; result.Removed
// is empty. The agent's mental model doesn't change — the EditResult
// shape is the same, just different fields populated.
```

### Undo after a bad call

```csharp
session.ReplaceText(anchor, /* something wrong */ "");
// Agent realizes the mistake or the user rejects it.
session.Undo();
// State is byte-equal to pre-op. Redo() would re-apply.
```

## Performance budgets (targets, not gates)

| Op | Target on a 100-page DOCX |
|---|---|
| `new DocxSession(bytes)` | < 250 ms |
| `ReplaceText` (1 paragraph) | < 5 ms + < 30 ms re-projection |
| `InsertParagraph`, `SplitParagraph` | < 5 ms + < 30 ms |
| `Project()` (full) | reuses converter budget: < 1 s |
| `Save()` | < 200 ms |
| `Undo()` | < 50 ms |
| Memory at 50-deep undo on a 5 MB DOCX | < 80 MB |

These are aspirations. Microbenchmarks aren't in CI by default — flag in PR if you measure 2× above target.

## Render-plan endpoints (the editor's incremental repaint)

Added for the browser editor's incremental structural repaint; useful to any renderer
that diffs a view against the session.

- **`ListBlocks()` → `RenderPlan { Body, Footnotes, Endnotes }`** — the ordered
  top-level render units per scope container: each body `w:p` under its projected kind
  (`p`/`h`/`li`, via the projector's `KindFor`), each `w:tbl` as ONE `tbl` unit (its
  rows/cells/cell paragraphs subsumed), and note definitions mirroring **exactly what
  the HTML renderer's notes section shows** (with ≥1 citation: the cited notes in
  citation order; with none: every non-separator note in part order — a rendered-but-
  uncited `continuationNotice` behaves differently in the two cases). Every
  `RenderUnit` carries `Sig`, a content hash (`UnidHelper.ContentHash`): in-session an
  element keeps its Unid across edits (ops rebuild children, not the block element),
  so unid alone cannot reveal an undone text edit or a row insert. Note ids —
  reference AND definition — are excluded from the hash: a footnote insert shifts
  every later note's `w:id` with no rendered-content change (marker/list numbering is
  position-derived chrome).
- **`ListNotes(endnotes = false)` → `IReadOnlyList<NoteListEntry { Id, DefAnchorId,
  Ordinal }>`** — footnotes/endnotes in citation order. The k-th marker in document
  order IS note k (ids ascend in reference order), so a client renumbers rendered note
  chrome (marker `sup` text, hrefs, `li` ids/values, backrefs) positionally instead of
  re-rendering every citing block.
- **`DocxSessionSettings.EmitMarkdownPatch`** (default `true`; wire
  `emitMarkdownPatch`) — when `false`, mutation ops return `Patch = null` and skip the
  per-op whole-document re-projection that builds it. Clients that re-render from HTML
  (the editor) should turn it off.
- **WASM-only companions** (`DocxSessionBridge`; not in the stdio host, same as
  `RenderHtml`/`RenderBlockHtml`): `ListAnchors` (the `{anchorIndex}` object without
  the markdown payload — the editor's per-op anchor-map refresh) and
  `RenderBlocksHtml(handle, anchorIdsJson, cssPrefix, fabricateClasses)` (batch block
  render: one throwaway document per call, each target cloned with its real siblings so
  `w:contextualSpacing` resolves, live `ListItemRetriever` annotations transplanted so
  an isolated list item renders its true number; a table returns with its generated
  alignment-`div` wrapper; `fn:`/`en:` anchors return their note paragraphs
  concatenated; a per-anchor failure maps to JSON `null`).

## Known limits and open questions

- **`MarkdownPatch.Markdown` is currently the full re-projection.** The `ScopeAnchorId` field correctly identifies the smallest enclosing block, but the payload is the whole document re-projected. A future optimization (per the spec's open questions) is to emit only the markdown for the named scope. Cheap mitigation: callers that care can splice using their cached projection.
- **Snapshot granularity is per-part XML clone.** For documents with very large embedded images or huge tables, per-element diffs would be more memory-efficient. Deferred until measured to be a problem.
- **Closing a session mid-flight from JS.** The WASM bridge holds sessions in a static dictionary keyed by handle; if a JS caller drops a `DocxSession` without calling `close()`, the .NET-side session is not eligible for GC. The npm wrapper exposes `Symbol.dispose` for TypeScript 5.2+ `using` blocks; older runtimes need explicit `.close()`.
- **`Save()` strips internal `PtOpenXml:Unid` attributes by default.** The projector assigns a Unid to every descendant of every projected scope; persisting them grows large documents by hundreds of KB of attribute noise (a 148 KB NVCA Model COI round-tripped at 588 KB before this default flipped). Anchor ids therefore do **not** survive `Save` → re-open by default — a fresh session re-assigns Unids and gets new ids. Set `DocxSessionSettings.PersistAnchorIds = true` to keep the ids (which keeps the bloat). This resolves Open Question #1 in `markdown_projection.md` in favor of "clean OOXML out by default, opt in to anchor stability."

## Related

- [`markdown_projection.md`](markdown_projection.md) — the read-side projector this builds on (anchor scheme, scope semantics, projector handlers)
- [`docx_converter.md`](docx_converter.md) — `WmlToHtmlConverter` internals (sibling write-side converter with very different goals)
- [`tracked_changes.md`](tracked_changes.md) — informs the `TrackedChangeMode` setting
- [`incremental_annotation_overlay.md`](incremental_annotation_overlay.md) — anchor-based overlay pattern; the read-side analog of this write-side API

## Inspection: document structure and formatting

`GetBlockMetadata` / `GetBlockMetadatas` / `GetListMembership` /
`GetSectionInfo` / `ListStyles` / `GetFormatting` / `ListInlineSpans` are pure reads — no
mutation, no undo snapshot, no projection invalidation. Each returns immutable records (or null
for an unknown/inapplicable single-anchor query).

### Styles and direct/effective formatting

`ListStyles()` enumerates the document's explicit paragraph, character, table, and numbering style
definitions. Each `StyleInfo` includes `Id`, `Name`, `Type`, `BasedOn`, `Next`, default/custom
flags, resolved latent-style gallery metadata, and the high-signal resolved paragraph/run/table
properties appropriate to its type. Resolution reuses `FormattingAssembler`'s style rollups — but
see **What "effective" includes** below: it is a *shorter cascade* than the renderer applies, not
the same one. A returned paragraph style `Id` is accepted unchanged by `SetParagraphStyle`; a
returned character style `Id` is accepted as `FormatOp.RunStyle`.

`GetFormatting(anchor)` is paragraph-only and explicitly separates:

- `DirectParagraph`: only properties present in that paragraph's `w:pPr`; absent values stay null.
- `EffectiveParagraph`: document defaults + full paragraph style chain + direct properties, with
  ordinary schema defaults filled for alignment, spacing, indentation, line spacing, and toggles.
- `Runs`: the same entries returned by `ListInlineSpans(anchor)`.

#### What "effective" includes — and what it does not

This is the resolver's exact contract. **It is not the render oracle's cascade**, and where the two
differ the render (`WmlToHtmlConverter`) is what Word actually shows.

`EffectiveParagraph` = `w:docDefaults/w:pPrDefault` + the `w:pStyle` `basedOn` chain + the
paragraph's direct `w:pPr`. It does **not** include:

- the **numbering level's** `w:pPr` (`w:abstractNum/w:lvl/w:pPr`, and a `w:lvlOverride/w:lvl`
  form of it), which the render path applies in `AssembleParagraphProperties` with its own
  `FromParagraph`/`FromStyle` priority;
- the **table style / `w:tblStylePr`** `w:pPr` layer for a paragraph inside a table.

`InlineSpan.Effective` = `w:docDefaults/w:rPrDefault` + the character/paragraph style chain +
the run's direct `w:rPr` + theme-font resolution. It does **not** toggle-merge the **table
style's** conditional `w:rPr` the way `AnnotateRunProperties` does.

Concretely, and pinned by `BM021`/`BM022`:

| Document | Renders as | `GetFormatting` reports |
|---|---|---|
| List item whose indent lives only in `w:abstractNum/w:lvl/w:pPr/w:ind` (the normal case) | indented | `LeftIndentTwips = 0` |
| Run in a `firstRow`-styled table whose bold comes from `w:tblStylePr` | bold | `Bold = false` |

`GetListMembership` surfaces the numbering level's real `Start`, `LevelText`, and indentation
separately, so the list case is recoverable by the caller today.

The exclusions are deliberate rather than accidental. The numbering layer is only applied by
`ParagraphStyleRollup` when the paragraph carries `ListItemRetriever`'s `ListItemInfo` annotation,
and that annotation is a lazily built **cache** — present or absent depending only on whether
something rendered, projected, or resolved a list label earlier in the session. Reading it would
make a pure read API answer the same unmutated document two different ways depending on call
order, so the resolver deliberately resolves against an annotation-free probe. That determinism
costs the numbering layer even for a projected session, which before this was the one case that
happened to get the fuller answer. Unifying the two cascades means changing the render oracle's
formatting path and is tracked separately.

Each `InlineSpan` reports the containing mutation-ready `AnchorId`, stable run `RunUnid`, flat-text
`Span`, text, `Direct` run properties, and `Effective` run properties. `AnchorId` + `Span` can be
passed directly to `ApplyFormat`. These are run/format spans only; hyperlink, bookmark, revision,
content-control, and other inline memberships are separate follow-ons (#451/#452/#455).

### `BlockMetadata`

For every block-level anchor, exposes:

- `AnchorId`, `Kind`, `Scope` — duplicated from `AnchorInfo` so the
  record is self-contained.
- `StyleId` / `StyleName` — `pStyle/@val` for paragraph kinds,
  `tblStyle/@val` for tables. `StyleName` resolves through the styles
  part's `w:name/@val`.
- `OutlineLevel` — `pPr/outlineLvl` when present; otherwise inferred
  from a `HeadingN` style (level 0..8). 0-based per Word convention.
- `List` — populated for list-item paragraphs (`null` otherwise).
- `HasInlineFormatting` — true when any descendant `w:r` carries a
  non-empty `w:rPr`. Coarse "does this paragraph have any character
  formatting at all" probe.

### `ListMembership`

For list-item paragraphs (and also surfaced as `BlockMetadata.List`):

- `NumId` / `AbstractNumId` / `Level` / `Format` — the standard
  numbering identity quadruple.
- `AnchorId` — the queried paragraph anchor, accepted unchanged by list mutations.
- `Start` / `LevelText` — the abstract level definition's start and marker template.
- `LeftIndentTwips` / `RightIndentTwips` / `FirstLineIndentTwips` /
  `HangingIndentTwips` — the effective level indentation (including a `w:lvlOverride/w:lvl`).
- `StartOverride` — non-null when the paragraph's `w:num` has a
  `w:lvlOverride/w:startOverride` at this level. Useful for predicting
  what `RestartNumberedList` will produce.
- `IsAutoNumbered` — always true (a paragraph without numbering returns
  `null` from `GetListMembership`).
- `FromStyle` — true when `w:numPr` is inherited from the paragraph's
  style chain (style → basedOn → basedOn → ...) rather than set inline.
  Lets callers reason about whether modifying the paragraph in place
  versus modifying the underlying style is appropriate.
- `GeneratedLabel` — same string as `AnchorInfo.AutoNumberPrefix`,
  duplicated here so callers don't take two round-trips.

### `SectionInfo`

For anchors in the body part:

- `AnchorId` — the queried body anchor, accepted unchanged by section mutations.
- `SectionUnid` — stable, stored Unid for the governing `w:sectPr` (never a positional fallback).
- `PageWidthTwips` / `PageHeightTwips` — raw twips (1 inch = 1440 twips).
- `Landscape` — true when `pgSz/@orient = "landscape"`.
- `MarginTopTwips` / `MarginBottomTwips` / `MarginLeftTwips` /
  `MarginRightTwips` — `pgMar` attribute values; defaults to 1440
  (1 inch) when missing.
- `Columns` — `cols/@num`, defaults to 1.
- `HeaderPartUris` / `FooterPartUris` — package-part URIs of the
  header/footer parts referenced via `headerReference` / `footerReference`,
  in declaration order. Empty when no headers/footers are referenced.
- `HeaderRefs` / `FooterRefs` — effective per-kind stories, including inherited references.
- `PageNumberStart` / `PageNumberFormat` — the section's explicit page-number settings.

Returns `null` for anchors in non-body parts (footnotes, endnotes,
headers, footers, comments) — sectPr is body-only.

### `NumberFormat` enum

Closed enum used by `ListMembership.Format` (read) and by the list
write surface (when it ships). Values: `Decimal`, `UpperLetter`,
`LowerLetter`, `UpperRoman`, `LowerRoman`, `Bullet`. Any OOXML
`numFmt` value outside this set maps to `Decimal` (safest fallback).
