# DOCX Mutation API Design

> **Status:** Approved design. Implementation will follow the nine-phase plan in [Implementation Phasing](#implementation-phasing). This document is the contract; the in-tree spec at `docs/architecture/docx_mutation_api.md` will be populated as phases land.

## Purpose

A stateful, in-memory editing API for Word documents, keyed by the anchor ids emitted by the markdown projection (see [`docs/architecture/markdown_projection.md`](../../architecture/markdown_projection.md)). The API is the substrate for agentic editing pipelines: an LLM agent reads the markdown projection, decides what to change, and invokes a small set of high-level mutation tools that take anchor ids and (mostly) markdown payloads — never raw OOXML, except through a deliberately separate escape hatch.

## Goals

1. **Anchor-keyed mutations.** Every edit names its target by an anchor id from the markdown projection. The agent never thinks in `XElement`s.
2. **Markdown in, markdown out.** Replacement and insertion payloads are written in the same markdown subset the projector emits, so the read/write surfaces stay symmetric.
3. **Stable session.** The document is loaded once and stays in memory across mutations. Each op runs in milliseconds; the agent doesn't pay re-parse cost per call.
4. **Predictable anchor lifecycle.** Every mutation returns exactly which anchors were created, removed, and modified — so the agent can keep its mental model in sync without re-projecting on every call.
5. **Typed errors, no exceptions across the boundary.** Every method returns a discriminated `EditResult`. Agents can pattern-match error codes; nothing throws into JS.
6. **Bounded undo.** A snapshot ring buffer lets agents recover from bad calls without orchestration.
7. **Raw escape hatch.** A clearly-namespaced `Raw.*` API lets callers inject arbitrary OOXML when the markdown subset can't express what they need (complex tables, math, content controls).
8. **Three-surface parity.** The .NET API, the WASM JSExport bridge, and the npm TypeScript wrapper ship together, same shape and same error codes.

## Non-Goals

- **Concurrent / collaborative editing.** One session, one writer. No CRDT, no OT, no locking. Agent harnesses own one session per logical document.
- **Markdown-to-DOCX conversion in general.** The payload parser handles a small, projector-symmetric subset. Full GFM and CommonMark are not in scope — for that, use `HtmlToWmlConverter` via a separate markdown→HTML step.
- **Structural table edits in v1** (insert/delete row, insert/delete column, merge cells). Cell content edits are in. Table structure is v2.
- **Direct image, footnote, or comment insertion in v1.** Stubbed with named v2 ops in the error catalog.
- **Persistence of the agent's edit log.** Operations are applied and snapshots are stored; we do not serialize the op stream itself. (Approach 2 in the design exploration would have enabled this; we deferred for v1.)

## Architecture

A new module sibling to `WmlToMarkdownConverter`, organized in three layers.

```
┌──────────────────────────────────────────────────────────────┐
│  npm: docxodus.openDocxSession(bytes) → DocxSession          │
│       session.replaceText(anchor, md) → EditResult           │
│       session.undo(), session.save() → Uint8Array            │
└──────────────────────────────────────────────────────────────┘
                            │  JS ↔ WASM (handle = int sessionId)
┌──────────────────────────────────────────────────────────────┐
│  wasm/DocxodusWasm/DocxSessionBridge.cs                      │
│    static Dictionary<int, DocxSession> _sessions;            │
│    [JSExport] OpenSession(byte[]) → int                      │
│    [JSExport] ReplaceText(int, string anchor, string md)     │
│       → string (JSON-serialized EditResult)                  │
└──────────────────────────────────────────────────────────────┘
                            │
┌──────────────────────────────────────────────────────────────┐
│  Docxodus/DocxSession.cs            (the real work)          │
│    sealed class DocxSession : IDisposable                    │
│      ctor(byte[], DocxSessionSettings)                       │
│      MarkdownProjection Project()                            │
│      EditResult ReplaceText(string anchor, string md)        │
│      EditResult DeleteBlock(string anchor)                   │
│      EditResult InsertParagraph(string anchor, Position, md) │
│      EditResult SplitParagraph(string anchor, int offset)    │
│      EditResult MergeParagraphs(string first, string second) │
│      EditResult ApplyFormat(string anchor, CharSpan?, FormatOp)│
│      EditResult SetParagraphStyle(string anchor, string)     │
│      EditResult SetListLevel(string anchor, int delta)       │
│      EditResult RemoveListMembership(string anchor)          │
│      EditResult ReplaceCellContent(string anchor, string md) │
│      RawDocxOps Raw { get; }                                 │
│      bool Undo();  bool Redo()                               │
│      byte[] Save()                                           │
└──────────────────────────────────────────────────────────────┘
                            │ owns
        ┌───────────────────┼────────────────────────┐
        ▼                   ▼                        ▼
  MemoryStream +      AnchorIndex            UndoRing
  WordprocessingDoc   (refreshed lazily)     (per-part XML snapshots,
  (long-lived)                                ring of N; default 50)
                            │
                            ▼
                  Internal: MarkdownPayloadParser
                  (markdown → list<Run> with rPr;
                   block markdown → list<XElement w:p>)
```

### Key choices

- **One live `WordprocessingDocument`.** Kept open over the session's `MemoryStream`. Mutations happen on the `XDocument` of the affected part (via the existing `GetXDocument()` extension). On `Save()`, `PutXDocument()` is called on dirtied parts and the stream bytes are copied out.
- **Reuse over reinvention.** Anchor resolution uses `AnchorTarget.Resolve`. Markdown rendering of the affected scope re-runs `WmlToMarkdownConverter` over just the dirtied scope. New code is the mutation surface, the inline-markdown parser, and the undo ring.
- **No run-level anchors.** Inline edits address `(paragraphAnchor, CharSpan)`. Runs split, merge, and re-form on every edit; addressing them creates a churn problem. Character spans within a stable paragraph are robust.
- **Markdown payload subset, not full CommonMark.** A small purpose-built parser handles exactly what the projector emits. Block: paragraphs, headings, lists, blockquotes, fenced code. Inline: bold, italic, code, strike, links, soft breaks, escapes. Everything else is typed-rejected with a remediation message.
- **Undo via per-part XML clones**, capped at `Settings.UndoDepth` (default 50). Most ops dirty only `MainDocumentPart`, so per-snapshot cost is small.
- **Anchor index updates incrementally.** Each op mutates the in-memory `AnchorIndex` (add/remove entries) rather than rebuilding. Full rebuild only after `Undo`/`Redo`.

## Public API

```csharp
#nullable enable

namespace Docxodus;

// ─── Session ────────────────────────────────────────────────────────────────

public sealed class DocxSession : IDisposable
{
    public DocxSession(byte[] docxBytes, DocxSessionSettings? settings = null);

    // View
    public MarkdownProjection Project();          // full re-projection (cached until next mutation)
    public bool Exists(string anchorId);
    public AnchorInfo? GetAnchorInfo(string anchorId);  // kind, scope, short text preview

    // Tier A — text CRUD
    public EditResult ReplaceText(string anchorId, string markdownPayload);
    public EditResult DeleteBlock(string anchorId);

    // Tier B — structural
    public EditResult InsertParagraph(string anchorId, Position pos, string markdownPayload);
    public EditResult SplitParagraph(string anchorId, int characterOffset);
    public EditResult MergeParagraphs(string firstAnchorId, string secondAnchorId);

    // Tier C — formatting
    public EditResult ApplyFormat(string anchorId, CharSpan? span, FormatOp op);
    public EditResult SetParagraphStyle(string anchorId, string styleId);
    public EditResult SetListLevel(string anchorId, int levelDelta);   // -1 outdent / +1 indent
    public EditResult RemoveListMembership(string anchorId);            // list item → normal paragraph

    // Tier D — table cell content (tracked-change mode is a session setting)
    public EditResult ReplaceCellContent(string cellAnchorId, string markdownPayload);

    // Raw OOXML escape hatch
    public RawDocxOps Raw { get; }

    // Lifecycle
    public bool Undo();
    public bool Redo();
    public byte[] Save();
    public void Dispose();

    // Diagnostics
    public Exception? LastInternalError { get; }   // most recent exception caught + rolled back
}

public sealed class RawDocxOps
{
    public string GetXml(string anchorId);
    public EditResult InsertXml(string anchorId, Position pos, string xml);
    public EditResult ReplaceXml(string anchorId, string xml);
}

// ─── Settings ───────────────────────────────────────────────────────────────

public sealed class DocxSessionSettings
{
    public int UndoDepth { get; init; } = 50;
    public bool ValidateRawOps { get; init; } = false;          // run OpenXmlValidator on raw inserts
    public TrackedChangeMode TrackedChanges { get; init; } = TrackedChangeMode.Accept;
    public string? RevisionAuthor { get; init; }                // used when TrackedChanges != Accept
    public WmlToMarkdownConverterSettings ProjectionSettings { get; init; } = new();
    public Microsoft.Extensions.Logging.ILogger? Logger { get; init; }   // optional; receives internal-error traces
}

// ─── Value types ────────────────────────────────────────────────────────────

public enum Position { Before, After }

public readonly record struct CharSpan(int Start, int Length);

public sealed record FormatOp
{
    public bool? Bold { get; init; }
    public bool? Italic { get; init; }
    public bool? Underline { get; init; }
    public bool? Strike { get; init; }
    public bool? Code { get; init; }                            // applies rStyle="Code"
    public string? Color { get; init; }                         // hex "FF0000" or "" to clear
    public string? RunStyle { get; init; }                      // arbitrary character style id
}
// null = leave as-is; true/false = set/clear explicitly

public sealed record AnchorInfo(string Id, string Kind, string Scope, string TextPreview);

// ─── Edit result ────────────────────────────────────────────────────────────

public sealed class EditResult
{
    public bool Success { get; init; }
    public EditError? Error { get; init; }                     // null iff Success
    public IReadOnlyList<Anchor> Created { get; init; } = Array.Empty<Anchor>();
    public IReadOnlyList<Anchor> Removed { get; init; } = Array.Empty<Anchor>();
    public IReadOnlyList<Anchor> Modified { get; init; } = Array.Empty<Anchor>();
    public MarkdownPatch? Patch { get; init; }                 // null on failure
}

public sealed record MarkdownPatch(string ScopeAnchorId, string Markdown);
// ScopeAnchorId names the enclosing block whose re-projection is in Markdown.
// For paragraph-level edits this is the paragraph's anchor; for inserts/deletes
// it is the nearest stable ancestor that still exists after the edit.

public sealed record EditError(EditErrorCode Code, string Message, string? AnchorId = null);

public enum EditErrorCode
{
    // Anchor problems
    AnchorNotFound, AnchorWrongKind, AnchorsNotAdjacent, SessionDisposed,

    // Markdown payload
    MalformedMarkdown, UnsupportedMarkdownSyntax,
    TableInsertNotSupported, FootnoteRefNotSupported,
    CommentMarkerNotSupported, ImageInsertNotSupported, AnchorTokenInPayload,

    // Positional / structural
    OffsetOutOfRange, InvalidPosition,

    // Formatting
    UnknownStyle, InvalidListLevel,

    // Raw OOXML
    MalformedXml, DisallowedNamespace, IncompatibleElementType, ValidationFailed,

    // Undo/redo
    NothingToUndo, NothingToRedo,

    // Catch-all for unexpected SDK exceptions (rollback applied)
    InternalError,
}
```

## Supported Markdown Subset

**Block payloads accepted by the parser** (mirror the projector):

| Syntax | Maps to |
|---|---|
| Blank-line separated text | `w:p` paragraphs |
| `#` … `######` | Paragraphs with style `Heading1` … `Heading6` |
| `- ` / `* ` / `+ ` | Bulleted list item (inherits the document's bulleted numbering definition at the insertion point) |
| `1. ` / `2. ` / … | Numbered list item (inherits ordered numbering definition) |
| Indentation under a list marker | Nested list level |
| `> ` | Paragraph styled `Quote` |
| Fenced code (```` ``` ````) or 4-space-indented | Paragraph styled `Code` |

**Inline payloads accepted:**

| Syntax | Maps to |
|---|---|
| `**bold**` | `w:r` with `w:b` |
| `*italic*` | `w:r` with `w:i` |
| `` `code` `` | `w:r` with `rStyle="Code"` |
| `~~strike~~` | `w:r` with `w:strike` |
| `[text](url)` | `w:hyperlink` |
| Soft break (single `\n`) | `w:br` |
| Backslash escapes | Literal character |

**Rejected with typed errors** (see [Error Catalog](#error-catalog) for codes and messages):

| Input | Reason |
|---|---|
| Pipe table | Use `ReplaceCellContent` for v1; `InsertTable` planned for v2 |
| `[^fn-id]` footnote/endnote reference | Output-only in v1; `AddFootnote` planned for v2 |
| `{#cmt:…}` inline comment marker | Output-only; `AddComment` planned for v2 |
| `![alt](docxodus://img/…)` | Needs binary upload; `AddImage` planned for v2 |
| Any `{#kind:scope:unid}` anchor token | Projection output, not input — strip from payload |
| Setext headings, HTML blocks, definition lists, other extensions | Outside the supported subset |

For anything markdown can never express (complex tables with merges, drawings, math, fields, content controls, custom XML), use `session.Raw.InsertXml` or `session.Raw.ReplaceXml` — see [Raw Escape Hatch](#raw-escape-hatch).

## Anchor Lifecycle

**Identity rule.** An anchor's identity is its `Unid` (the 32-char hex tail). The `kind:scope:` prefix is descriptive metadata and **may change across mutations** — promoting a `Normal` paragraph to `Heading2` flips its anchor id from `{#p:body:abcd}` to `{#h:body:abcd}`. `AnchorTarget.Resolve` walks by Unid alone, so old full-form ids continue to resolve. Agents that cache a full id should treat the prefix as stale-friendly and prefer the `Modified` entry in `EditResult` for the current canonical form.

**Per-op lifecycle** (with `TrackedChanges = Accept`; see [tracked-change footnote](#tracked-change-mode) for the alternative):

| Op | Created | Removed | Modified | Patch scope |
|---|---|---|---|---|
| `ReplaceText(p, md)` | (none in v1) | descendant inline anchors that no longer exist (rare) | `p` | `p` |
| `DeleteBlock(p)` | — | `p` + all descendant anchors | — | nearest stable ancestor (scope section or enclosing cell) |
| `InsertParagraph(p, pos, md)` | one anchor per new block | — | — | smallest enclosing common parent of `p` + new blocks |
| `SplitParagraph(p, offset)` | **second half** | — | `p` (first half — original anchor convention) | enclosing parent |
| `MergeParagraphs(first, second)` | — | `second` + descendant anchors | `first` | `first` |
| `ApplyFormat(p, span?, op)` | — | — | `p` | `p` |
| `SetParagraphStyle(p, style)` | — | — | `p` (potentially new `kind` prefix) | `p` |
| `SetListLevel(p, delta)` | — | — | `p` | enclosing list (downstream items renumber) |
| `RemoveListMembership(p)` | — | — | `p` (`kind` flips `li`→`p`) | enclosing list |
| `ReplaceCellContent(tc, md)` | — | descendant inline anchors (rare) | `tc` | `tc` |
| `Raw.InsertXml(a, pos, xml)` | every block in the new XML | — | — | enclosing parent |
| `Raw.ReplaceXml(a, xml)` | every block in the new XML | `a` + descendants | — | enclosing parent |
| `Raw.GetXml(a)` | — | — | — | — (read-only; not an `EditResult`) |
| `Undo()` / `Redo()` | (diff vs current) | (diff vs current) | (diff vs current) | `null` (caller should `Project()` for a fresh full view) |

**Conventions:**

- **`SplitParagraph`** — original Unid stays on the first half. External references (LLM context, search indexes) are biased toward the pre-split position, so keeping the prefix half stable minimizes invalidation.
- **`MergeParagraphs`** — first anchor absorbs the second. First is to the left in reading order, more likely to be the addressed one.
- **`InsertParagraph` with multi-paragraph payload** — anchors returned in document order (top to bottom in the new content).
- **`ReplaceText` is a content swap, not a paragraph rebuild.** The `w:p` element and its `w:pPr` (style, numbering, indentation) are preserved — only run children are replaced. A list item that gets `ReplaceText`'d stays a list item; a `Heading2` stays a `Heading2`.

### Tracked-change mode

When `Settings.TrackedChanges = RenderInline`, the "Removed" semantics shift: `DeleteBlock` wraps the paragraph in `w:del` rather than removing the element, so the anchor stays live (and appears in `Modified` instead of `Removed`). `ReplaceText` keeps the old runs wrapped in `w:del` and inserts new ones in `w:ins`. The agent sees `Modified` only; the document carries the full revision history. The `EditResult` shape is unchanged — just different fields populated.

### Failure semantics

On any error code, `Created`/`Removed`/`Modified` are empty, `Patch` is `null`, and **no mutation has been applied** — the snapshot taken at op-start was rolled back. Session state is identical to its pre-op state. This holds for both pre-apply validation failures and runtime failures caught by the rollback path. Failed ops do not consume an undo slot.

## Validation Pipeline

Every mutation flows through one pipeline:

```
1. Pre-checks (cheap, no state change)
   ├─ Session not disposed                     → SessionDisposed
   ├─ Anchor exists in current index           → AnchorNotFound
   ├─ Anchor's kind is valid for this op       → AnchorWrongKind
   ├─ Position/offset/span in range            → OffsetOutOfRange, InvalidPosition
   ├─ Style id exists in styles part           → UnknownStyle
   └─ For Raw ops: well-formed + namespace + slot-compatible
                                                → MalformedXml,
                                                  DisallowedNamespace,
                                                  IncompatibleElementType
2. Payload parse (markdown ops only)
   ├─ Parse markdown payload                   → MalformedMarkdown
   └─ Reject syntax outside the supported subset
                                                → UnsupportedMarkdownSyntax,
                                                  TableInsertNotSupported,
                                                  FootnoteRefNotSupported,
                                                  CommentMarkerNotSupported,
                                                  ImageInsertNotSupported,
                                                  AnchorTokenInPayload
3. Snapshot                          ← per-part XML clone pushed to undo ring
4. Apply                                       (no errors expected here; SDK
                                                exceptions caught and surfaced
                                                as InternalError + rollback)
5. Post-checks (optional, slower)
   └─ ValidateRawOps=true: re-run OpenXmlValidator on dirtied parts;
      if new errors appeared compared to pre-op baseline
                                                → ValidationFailed (rollback)
6. Incremental anchor index update + scope re-projection
                                                → never fails; the projector
                                                  is total over valid OOXML
7. EditResult { Success=true, Created/Removed/Modified, Patch }
```

## Error Catalog

| Group | Codes | Agent should… |
|---|---|---|
| Stale anchor | `AnchorNotFound` | Re-`Project()` and re-derive the anchor from current text |
| Wrong-shaped call | `AnchorWrongKind`, `AnchorsNotAdjacent`, `InvalidPosition`, `OffsetOutOfRange` | Re-read anchor kind via `GetAnchorInfo`; reissue with the right op or coordinates |
| Bad payload | `MalformedMarkdown`, `UnsupportedMarkdownSyntax`, `AnchorTokenInPayload` | Fix the markdown; the message names what was wrong |
| Use the right op | `TableInsertNotSupported`, `FootnoteRefNotSupported`, `CommentMarkerNotSupported`, `ImageInsertNotSupported` | Call the named v1 op (when one exists) or fall back to `Raw.InsertXml` |
| Bad style / list level | `UnknownStyle`, `InvalidListLevel` | Re-query (no `ListStyles()` API in v1; agent guesses from projection; v2 may add) |
| Bad raw XML | `MalformedXml`, `DisallowedNamespace`, `IncompatibleElementType`, `ValidationFailed` | Use `Raw.GetXml(anchor)` as a template; mutate; resubmit |
| Lifecycle | `SessionDisposed`, `NothingToUndo`, `NothingToRedo` | Stop / reopen / accept |
| Internal | `InternalError` | Should not happen — bug if it does. Op rolled back; safe to retry once or report. Full exception is logged via the session's `ILogger` (if provided) and accessible via `session.LastInternalError`. |

**Cross-boundary contract.** Every public method on `DocxSession` and `RawDocxOps` returns `EditResult` (or a read-only value for `Project`, `Exists`, `GetAnchorInfo`, `GetXml`). The constructor and `Save()` are the only places that can throw — and only for genuinely fatal conditions (invalid DOCX bytes on open, IO failure on save). Agents never need a try/catch around individual edits.

## Raw Escape Hatch

```csharp
session.Raw.GetXml(string anchor) → string
    // Returns the OOXML for the element, useful as a template the
    // agent edits and feeds back via ReplaceXml.

session.Raw.InsertXml(string anchor, Position pos, string xml) → EditResult
    // Insert a block-level XML fragment before/after the anchor.
    // For w:p, w:tbl, w:sdt at body level; new Unids auto-assigned.

session.Raw.ReplaceXml(string anchor, string xml) → EditResult
    // Replace the anchored element with the supplied fragment.
    // Element-type compatible swap (paragraph→paragraph or block; cell→cell).
```

Validation pipeline for `Raw.*` ops (cheap → expensive; short-circuits):

1. **Well-formedness** — `XElement.Parse`; reject `MalformedXml`.
2. **Namespace check** — root and descendants must be in an allowed list (`w:`, `m:` for math, `wp:`/`a:` for drawing, `r:` for relationships). Reject `DisallowedNamespace` otherwise — no script, no foreign XML.
3. **Structural slot check** — root element type must be legal at the insertion point. `w:p`/`w:tbl`/`w:sdt`/`w:sectPr` for body-level; `w:r`/`w:hyperlink`/`w:fldSimple`/`w:sdt` for run-level. Hand-maintained whitelist; reject `IncompatibleElementType`.
4. **Optional deep validation** — `Settings.ValidateRawOps = true` runs `OpenXmlValidator` on the affected part post-apply; rolls back via the snapshot if validation errors increased. Off by default (slow); recommended on for untrusted agent input.

Rollback on rejection is free — the snapshot is taken before the op runs, so a failed validation discards the snapshot without restoring.

**Round-trip pattern for agents:** `GetXml(anchor)` → edit a single attribute or child → `ReplaceXml(anchor, modified)`. Much safer than authoring from scratch.

**Progressive enhancement path:** every v2 typed op (`InsertTable`, `AddImage`, `AddComment`) will be implemented as a thin wrapper over `Raw.InsertXml` with a typed builder. v1 raw escape hatch + v2 typed builders is a clean ladder.

## WASM / npm Surface

### WASM bridge

Sessions live on the .NET heap and persist across `[JSExport]` calls, so the bridge needs a session registry keyed by an integer handle.

```csharp
// wasm/DocxodusWasm/DocxSessionBridge.cs
public static class DocxSessionBridge
{
    private static readonly Dictionary<int, DocxSession> _sessions = new();
    private static int _nextId = 1;

    [JSExport] public static int OpenSession(byte[] docxBytes, string settingsJson);
    [JSExport] public static void CloseSession(int handle);
    [JSExport] public static string Project(int handle);                              // JSON

    [JSExport] public static string ReplaceText(int h, string anchor, string md);     // JSON EditResult
    [JSExport] public static string DeleteBlock(int h, string anchor);
    [JSExport] public static string InsertParagraph(int h, string anchor, string posStr, string md);
    [JSExport] public static string SplitParagraph(int h, string anchor, int offset);
    [JSExport] public static string MergeParagraphs(int h, string first, string second);
    [JSExport] public static string ApplyFormat(int h, string anchor, string spanJson, string opJson);
    [JSExport] public static string SetParagraphStyle(int h, string anchor, string styleId);
    [JSExport] public static string SetListLevel(int h, string anchor, int delta);
    [JSExport] public static string RemoveListMembership(int h, string anchor);
    [JSExport] public static string ReplaceCellContent(int h, string anchor, string md);

    [JSExport] public static string RawGetXml(int h, string anchor);
    [JSExport] public static string RawInsertXml(int h, string anchor, string posStr, string xml);
    [JSExport] public static string RawReplaceXml(int h, string anchor, string xml);

    [JSExport] public static string Undo(int h);
    [JSExport] public static string Redo(int h);
    [JSExport] public static byte[] Save(int h);
}
```

Two non-obvious choices:

- **Strings, not opaque references**, across the boundary — every complex argument and every return is JSON. Keeps the bridge surface trivially serializable and avoids the .NET-WASM object identity headaches.
- **Explicit `CloseSession`** is mandatory. JS-side `using`-like semantics are wrong because `[JSExport]` methods can't observe JS GC. The npm wrapper hides this behind a `Symbol.dispose`/`FinalizationRegistry` shim so callers can use `using` blocks under TypeScript 5.2+ or rely on GC under older runtimes.

### npm wrapper

```typescript
// npm/src/index.ts
import { DocxSession } from './session.js';

export async function openDocxSession(
  bytes: Uint8Array,
  settings?: DocxSessionSettings
): Promise<DocxSession>;

// npm/src/session.ts
export class DocxSession implements Disposable {
  project(): Promise<MarkdownProjection>;
  replaceText(anchorId: string, markdown: string): Promise<EditResult>;
  deleteBlock(anchorId: string): Promise<EditResult>;
  insertParagraph(anchorId: string, pos: 'before' | 'after', md: string): Promise<EditResult>;
  splitParagraph(anchorId: string, offset: number): Promise<EditResult>;
  mergeParagraphs(first: string, second: string): Promise<EditResult>;
  applyFormat(anchorId: string, span: CharSpan | null, op: FormatOp): Promise<EditResult>;
  setParagraphStyle(anchorId: string, styleId: string): Promise<EditResult>;
  setListLevel(anchorId: string, delta: number): Promise<EditResult>;
  removeListMembership(anchorId: string): Promise<EditResult>;
  replaceCellContent(anchorId: string, markdown: string): Promise<EditResult>;

  readonly raw: {
    getXml(anchorId: string): Promise<string>;
    insertXml(anchorId: string, pos: 'before' | 'after', xml: string): Promise<EditResult>;
    replaceXml(anchorId: string, xml: string): Promise<EditResult>;
  };

  undo(): Promise<boolean>;
  redo(): Promise<boolean>;
  save(): Promise<Uint8Array>;
  close(): void;
  [Symbol.dispose](): void;
}
```

```typescript
type EditResult = {
  success: boolean;
  error?: { code: EditErrorCode; message: string; anchorId?: string };
  created: Anchor[];
  removed: Anchor[];
  modified: Anchor[];
  patch?: { scopeAnchorId: string; markdown: string };
};

type EditErrorCode =
  | 'anchor_not_found' | 'anchor_wrong_kind' | 'anchors_not_adjacent'
  | 'session_disposed'
  | 'malformed_markdown' | 'unsupported_markdown_syntax'
  | 'table_insert_not_supported' | 'footnote_ref_not_supported'
  | 'comment_marker_not_supported' | 'image_insert_not_supported'
  | 'anchor_token_in_payload'
  | 'offset_out_of_range' | 'invalid_position'
  | 'unknown_style' | 'invalid_list_level'
  | 'malformed_xml' | 'disallowed_namespace' | 'incompatible_element_type'
  | 'validation_failed'
  | 'nothing_to_undo' | 'nothing_to_redo'
  | 'internal_error';
```

`EditErrorCode` is generated from the C# enum at build time via a small codegen step in `npm/scripts/` so the two surfaces stay in lock-step. Agents pattern-match on snake_case strings.

All session work runs on the existing `docxodus.worker.ts` web worker so the main thread stays responsive. The worker proxy already exists for `convertWmlToMarkdown`; we extend it with a session-handle-aware message protocol (`{ kind: 'session.op', handle, op, args }`).

### React hook

```typescript
// npm/src/react.ts
export function useDocxSession(bytes: Uint8Array | null): {
  session: DocxSession | null;
  loading: boolean;
  error: Error | null;
};
```

Auto-disposes on unmount. Matches the existing `useDocxToMarkdown` hook shape.

## Testing Strategy

**File layout.** New test class `Docxodus.Tests/DocxSessionTests.cs`, mirroring the existing prefix convention. Test IDs prefixed `DS###`. Fixtures reuse `TestFiles/HC*` (heading-heavy legal docs already used by the markdown projection) and add a small set under `TestFiles/DS*` for table cells, lists, and tracked-change scenarios.

| Category | Coverage | Count | Notes |
|---|---|---|---|
| Per-op happy path | Every public method × representative fixture | ~15 | Open session → call op → assert `Success=true` and expected `Created/Removed/Modified` shape |
| Per-op failure path | Every `EditErrorCode` triggered at least once | ~22 | One test per code; assert no state change (snapshot byte-equal pre/post) |
| Anchor lifecycle | The lifecycle table, verified empirically | ~12 | For each op: assert returned diff matches the policy; removed anchors no longer resolve; created anchors resolve and round-trip via `Resolve` |
| Markdown payload parser | Inline subset + block subset + rejections | ~20 | Pure parser tests (parser is unit-testable independent of session) |
| Patch correctness | After op, splice `Patch.Markdown` into a cached projection at `ScopeAnchorId`; assert equal to fresh `Project()` | ~8 | One per scope variant (paragraph / list / cell / multi-paragraph insert / delete) |
| Undo / redo | Full op coverage × `Undo` × `Redo` | ~6 | Capture pre-op state, run op, undo, assert equal; redo, assert post-op state; chained ops + depth limit |
| Round-trip integrity | After arbitrary edit sequence: `Save()` → re-open → `Project()` → matches the session's last `Project()` | ~5 | Catches "edits not persisted to underlying part" bugs |
| Tracked-change mode | Each mutating op under `TrackedChanges = RenderInline`; assert `w:ins`/`w:del` markup and `Removed` stays empty | ~10 | Critical because the lifecycle footnote changes semantics |
| Raw escape hatch | `GetXml` → mutate → `ReplaceXml` round trip; namespace rejection; slot incompatibility; validation-on with deliberate bad XML | ~8 | `ValidateRawOps=true` path verified end-to-end |
| Property tests | Random op sequences on a small fixture; assert: anchor index internally consistent, undo-all returns to original bytes (modulo whitespace), no orphaned Unids | ~3 fixtures × ≥100 sequences | xUnit + a tiny in-test generator; no FsCheck dependency |

**Concurrency policy.** `DocxSession` is **not thread-safe** for mutations; concurrent reads (`Project`, `GetAnchorInfo`, `Exists`) on a quiescent session are safe. No locks — agent harnesses own one session per logical document. One test asserts the documented behavior (concurrent mutate races throw or corrupt — we don't make it work, just confirm it's not silently broken).

**WASM/Playwright.** New spec `npm/tests/docx-session.spec.ts` mirroring `npm/tests/markdown.spec.ts`:

- Session lifecycle: `openSession(bytes) → handle → mutate → save → re-open → project equals last`
- Each tier represented by one or two ops (full op-set coverage stays in .NET; npm tests prove the bridge works, not the logic)
- Error propagation: trigger one error from each group and assert `success === false` with the right `error.code`
- Undo across the bridge
- Memory: open and dispose 50 sessions in a loop, assert WASM heap doesn't grow unboundedly (catches the "static `_sessions` dictionary leaks" failure mode)

### Fixture additions

- `TestFiles/DS001_simple_two_paragraphs.docx` — minimal smoke fixture
- `TestFiles/DS002_lists_nested.docx` — bulleted + numbered, nested
- `TestFiles/DS003_table_with_cells.docx` — for `ReplaceCellContent`
- `TestFiles/DS004_tracked_changes_seed.docx` — pre-existing tracked changes
- `TestFiles/DS005_raw_complex.docx` — contains an SDT and a math equation

## Performance Budgets

Targets, not hard gates — flagged in PR if exceeded by >2×.

| Operation | Target (100-page DOCX) |
|---|---|
| `new DocxSession(bytes)` | < 250 ms |
| `ReplaceText` (1 paragraph) | < 5 ms (op apply) + < 30 ms (scope re-projection) |
| `InsertParagraph` (1 paragraph) | < 5 ms + < 30 ms |
| `SplitParagraph` | < 5 ms + < 30 ms |
| `Project()` (full) | reuses converter budget: < 1 s |
| `Save()` | < 200 ms |
| `Undo()` | < 50 ms (snapshot swap + index rebuild) |
| Memory per session at 50-deep undo ring on a 5 MB DOCX | < 80 MB |

Microbenchmarks live in `Docxodus.Tests/Benchmarks/DocxSessionBench.cs` as ordinary xUnit tests with `[Trait("Category", "Performance")]` so they're filterable; not run by default in CI.

## Implementation Phasing

Each phase is independently mergeable; each ends green with its own tests. Same staging discipline as the markdown projection's eight-phase rollout.

| Phase | Scope | New code | Tests | Mergeable as |
|---|---|---|---|---|
| **1. Skeleton** | `DocxSession` type, settings, value types, `Project`, `Save`, `Exists`, `GetAnchorInfo`, internal `UndoRing` infra (no public undo yet) | ~600 LOC core + ~200 LOC tests | DS001–DS005 (open/project/save round-trip; anchor info) | `feat(session): skeleton + projection passthrough` |
| **2. Markdown payload parser** | `Internal/MarkdownPayloadParser.cs` — block + inline subset, rejection codes | ~500 LOC + ~400 LOC tests | DS010–DS030 (parser tests — pure, no session needed) | `feat(session): markdown payload parser` |
| **3. Tier A — text CRUD** | `ReplaceText`, `DeleteBlock`. First wiring of pipeline (snapshot → apply → reproject → patch). Public `Undo`/`Redo`. | ~400 LOC + ~300 LOC tests | DS040 happy/sad + lifecycle + undo coverage | `feat(session): tier A text CRUD + undo` |
| **4. Tier B — structural** | `InsertParagraph`, `SplitParagraph`, `MergeParagraphs` | ~500 LOC + ~300 LOC tests | DS060 — split anchor convention, merge adjacency check, multi-paragraph insert | `feat(session): tier B structural ops` |
| **5. Tier C — formatting** | `ApplyFormat`, `SetParagraphStyle`, `SetListLevel`, `RemoveListMembership`. Run-splitting helper for span formatting | ~600 LOC + ~300 LOC tests | DS080 — span formatting splits/merges runs correctly, style change flips anchor kind, list level renumbers | `feat(session): tier C formatting + styles` |
| **6. Tier D + tracked changes** | `ReplaceCellContent`. `TrackedChanges = RenderInline` path applied retroactively across all prior ops. | ~300 LOC + ~400 LOC tests | DS100 — cell ops + tracked variants of every earlier op | `feat(session): tier D + tracked-change mode` |
| **7. Raw escape hatch** | `RawDocxOps` + `Settings.ValidateRawOps` | ~400 LOC + ~300 LOC tests | DS120 — GetXml/InsertXml/ReplaceXml + validation rejections | `feat(session): raw OOXML escape hatch` |
| **8. WASM + npm + React + Playwright** | `DocxSessionBridge.cs`, npm `session.ts`, worker proxy extension, React hook, codegen for `EditErrorCode`, Playwright suite | ~800 LOC + ~500 LOC tests | `npm/tests/docx-session.spec.ts` end-to-end | `feat(session): WASM bridge + npm/React wrappers` |
| **9. Docs & changelog** | `docs/architecture/docx_mutation_api.md` (in-tree spec, expanded with worked examples) + CHANGELOG `[Unreleased]` entry + CLAUDE.md module entry under "Core Modules" | docs only | — | `docs(session): architecture + changelog` |

Rough total: ~4,200 LOC code + ~3,000 LOC tests across nine PRs. Comparable to the markdown projection's eight-phase rollout (~5,000 LOC).

## Worked Examples

### Replace a clause

```typescript
const session = await openDocxSession(bytes);
const proj = await session.project();
// agent reads markdown, finds the indemnification clause anchor
const anchor = 'p:body:a1b2c3d4e5f6...';

const result = await session.replaceText(
  anchor,
  'The **Provider** shall indemnify the *Client* for any [breach](https://example.com/terms#breach) of the foregoing.'
);

// result.success === true
// result.modified === [{ id: 'p:body:a1b2...', kind: 'p', scope: 'body', unid: 'a1b2...' }]
// result.patch.scopeAnchorId === 'p:body:a1b2...'
// result.patch.markdown === '{#p:body:a1b2c3d4e5f6...} The **Provider** shall indemnify...'
```

### Split a paragraph and promote the second half to a heading

```typescript
const split = await session.splitParagraph(anchor, 42);
// split.modified[0] === original anchor (first half)
// split.created[0] === new anchor for the second half
const secondHalf = split.created[0].id;

await session.setParagraphStyle(secondHalf, 'Heading2');
// the secondHalf anchor's kind prefix is now 'h' instead of 'p';
// resolution by Unid is unaffected
```

### Inject a structured content control via raw XML

```typescript
const xml = await session.raw.getXml('p:body:abcd...');
// agent modifies the XML to wrap the paragraph in an <w:sdt> for structured tagging
const modified = wrapInSdt(xml, { tag: 'PartyName', alias: 'Party Name' });
const result = await session.raw.replaceXml('p:body:abcd...', modified);
// result.created includes the new SDT anchor and the (preserved) paragraph anchor
```

### Tracked-change mode

```typescript
const session = await openDocxSession(bytes, {
  trackedChanges: 'render_inline',
  revisionAuthor: 'agent-alpha',
});

await session.replaceText(anchor, 'Updated clause text.');
// Document now contains <w:del> wrapping old runs and <w:ins> wrapping new runs.
// The paragraph anchor stays live; result.removed is empty.
```

## Open Questions

- **Markdown parser library vs. hand-roll.** A purpose-built parser for the projector's subset is small (~500 LOC) and avoids a dependency. Alternative: pull in `Markdig` and post-process its AST. Hand-roll is the default; revisit if the subset grows.
- **List numbering inheritance on `InsertParagraph` at list boundaries.** When inserting `- Item` adjacent to an existing bulleted list, do we join the existing list's numbering definition, or create a new one? Default: join if `pos === 'after'` and `anchor` is `li`; create new otherwise. May need a `listJoin: 'auto' | 'join' | 'new'` parameter in v1.1.
- **Snapshot granularity.** Per-part XML clones are the v1 plan. If undo memory becomes a problem (large embedded images, huge tables), a future optimization is per-element diffs. Defer until measured.
- **Closing a session mid-flight.** What happens to in-progress worker messages when the agent calls `close()` between an op and its result? Plan: worker rejects pending messages for closed handles with a `SessionClosed` JS error (distinct from the `SessionDisposed` C# error code, which is only emitted when a method is called on a disposed session from C#).
- **`Save()` between mutations vs. only at end.** Should `Save()` invalidate the undo ring? Plan: no — saving is a snapshot of bytes for external use; the in-memory session continues with full history.

## Related

- [`markdown_projection.md`](../../architecture/markdown_projection.md) — the read-side projector this builds on
- [`docx_converter.md`](../../architecture/docx_converter.md) — `WmlToHtmlConverter` internals
- [`tracked_changes.md`](../../architecture/tracked_changes.md) — informs the `TrackedChangeMode` setting
- [`incremental_annotation_overlay.md`](../../architecture/incremental_annotation_overlay.md) — anchor-based overlay pattern; the read-side analog of this write-side API
