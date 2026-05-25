# DOCX Mutation API

> **Status:** Implemented. Source: `Docxodus/DocxSession.cs`, `Docxodus/RawDocxOps.cs`, `Docxodus/Internal/MarkdownPayloadParser.cs`, `Docxodus/Internal/UndoRing.cs`. Tests: `Docxodus.Tests/DocxSessionTests.cs` (`DS###`), `MarkdownPayloadParserTests.cs`, and an end-to-end smoke at `DocxSessionSmokeTest.cs`. WASM bridge: `wasm/DocxodusWasm/DocxSessionBridge.cs`. npm wrapper: `npm/src/session.ts`. The full type-level spec lives at `docs/superpowers/specs/2026-05-24-docx-mutation-api-design.md` — this doc is the conceptual reading and the recipe book; it points to source for the canonical shapes rather than restating them.

## What this is

`DocxSession` is the **write-side counterpart** to `WmlToMarkdownConverter`. The projector turns a DOCX into anchor-addressed markdown; the session lets you mutate the same DOCX by those anchor ids — replace text, insert/split/merge paragraphs, apply formatting, edit table cells — without the agent (or human) ever having to think about OOXML. Anything the markdown subset can't express drops to a clearly-namespaced raw-XML escape hatch.

The intended consumer is an agentic editing pipeline: an LLM reads the markdown projection of a document, decides what to change, and calls a small set of high-level tools. But the same surface is useful for any tooling that wants to make surgical, ID-addressed edits to Word documents — review pipelines, structured-edit UIs, templating workflows.

## Why it's shaped the way it is

Three design forces, in order of weight:

**The agent must not learn OOXML.** Every public method takes an anchor id (a string) and either a markdown payload (a string) or a small typed value (a `FormatOp`, a `CharSpan`). The agent never sees an `XElement`, never picks an SDK type, never has to know that bold is `w:b` inside `w:rPr`. The Raw escape hatch exists for the cases the markdown subset can't reach, but it's a separate namespace (`session.Raw.*`) so it's syntactically obvious when you've left the safe zone.

**Edits must be reversible.** Agents make mistakes. The session keeps a bounded ring of pre-op snapshots (default 50 deep) so `Undo()` and `Redo()` work without the caller orchestrating anything. Snapshots are per-part XML clones, not full package round-trips, so the cost is proportional to the size of the part the op touched — usually just the body.

**Errors must be pattern-matchable, not stringly-typed.** Every mutation returns an `EditResult` envelope; failure carries a typed `EditErrorCode` with a remediation message. The same enum is exposed as a snake-case string union in TypeScript, so JS agents pattern-match the same way C# callers do. No method on the session throws across the boundary (the constructor and `Save()` are the only places that can — and only for fatal conditions like an invalid DOCX or IO failure).

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

When you pass markdown into `ReplaceText`, `InsertParagraph`, or `ReplaceCellContent`, the session runs it through `MarkdownPayloadParser`, a hand-rolled parser that accepts **only** what the projector emits. Block-level: paragraphs, ATX headings (`#`–`######`), bulleted lists, ordered lists (with indent-based nesting), blockquotes, fenced code blocks. Inline: `**bold**`, `*italic*`, `` `code` ``, `~~strike~~`, `[text](url)` links, soft breaks, backslash escapes.

This is symmetric by design: anything the projector can emit, the parser can accept, so an agent can read markdown out and write markdown in. Anything outside the subset is rejected with a typed error that names either the v1 op to use instead or the v2 op planned to address it. The full table of accepted and rejected syntax is in the spec — the practical shorthand:

- If you can see it in the projection output, you can write it in a payload.
- If you need a pipe table → use `ReplaceCellContent` on each cell (`InsertTable` is v2).
- If you need a footnote, comment, or image → those are v2 ops, currently rejected with a clear error.
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
| `DeleteBlock(p)` | — | `p` + all descendant anchors | — | nearest stable ancestor |
| `InsertParagraph(p, pos, md)` | one anchor per new block | — | — | smallest enclosing common parent |
| `SplitParagraph(p, offset)` | the **second** half | — | `p` (first half — convention) | enclosing parent |
| `MergeParagraphs(a, b)` | — | `b` + descendants | `a` | `a` |
| `ApplyFormat(p, span?, op)` | — | — | `p` | `p` |
| `SetParagraphStyle(p, style)` | — | — | `p` (kind prefix may flip) | `p` |
| `SetListLevel(p, delta)` | — | — | `p` | enclosing list (downstream items renumber) |
| `RemoveListMembership(p)` | — | — | `p` (kind flips `li`→`p`) | enclosing list |
| `ReplaceCellContent(tc, md)` | — | descendant inline anchors (rare) | `tc` | `tc` |
| `Raw.InsertXml(a, pos, xml)` | every block in the new XML | — | — | enclosing parent |
| `Raw.ReplaceXml(a, xml)` | unids present in the new XML but not the old (typical for caller-authored XML) | unids present in the old element but not the new (when `a` itself is gone) | unids present in both (typical for the `GetXml → mutate → ReplaceXml` round trip, which preserves Unids) | enclosing parent |
| `Undo()` / `Redo()` | (diff vs current) | (diff vs current) | (diff vs current) | `null` — caller re-projects |

Two conventions worth pinning down because they affect agent reasoning:

- **`SplitParagraph` keeps the original Unid on the first half.** Reason: external systems (LLM context windows, search indices) bias toward the pre-split anchor position; keeping the prefix-half stable minimizes invalidation downstream.
- **`MergeParagraphs` lets the first anchor absorb the second.** Symmetric reason: the first anchor is to the left in reading order and is more likely to be the one a caller has cached.

**Tracked-change mode shifts the semantics for `ReplaceText` and `DeleteBlock`.** When `Settings.TrackedChanges = RenderInline`, deletions don't remove elements — they wrap old runs in `w:del` and new content in `w:ins`. So the affected anchor stays live and appears in `Modified` instead of `Removed`. The agent's view of the world doesn't have to change; the `EditResult` shape is unchanged.

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
├── Just the bold/italic/underline/code/color of some characters?
│       → ApplyFormat(anchor, CharSpan(start, length), FormatOp{...})
│       → ApplyFormat(anchor, null, FormatOp{...})  # null span = whole paragraph
│
├── Changing a paragraph's style (e.g., Normal → Heading2)?
│       → SetParagraphStyle(anchor, styleId)
│
├── Indenting/outdenting a list item or removing it from a list?
│       → SetListLevel(anchor, +1 | -1)
│       → RemoveListMembership(anchor)
│
├── Replacing the contents of a table cell?
│       → ReplaceCellContent(tcAnchor, markdown)
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

## Finding anchors via tagged annotations

The session addresses content by anchor id, but real workflows don't start with anchor ids — they start with intent ("edit the indemnification provision," "tighten the termination clause"). The clean way to bridge intent to anchors is to **annotate the regions ahead of time**, then resolve the annotation to its anchor(s) at edit time.

Docxodus's `AnnotationManager` already persists annotations into the docx itself: each annotation creates a `w:bookmark` named `_Docxodus_Ann_<id>` covering the range, and a custom XML part stores the metadata (`LabelId`, `Label`, `Color`, `Metadata` key/value bag). See [`custom_annotations.md`](custom_annotations.md) for the full mechanism and lifecycle. Annotations survive save/reopen and travel with the document.

A first-class `DocxSession.FindByAnnotation(id)` / `FindByLabel(labelId)` API is planned (tracked in [#132](https://github.com/JSv4/Docxodus/issues/132)). Until that lands, the bridge is straightforward enough to do in two steps with the existing surface:

```csharp
// Step 1 — read annotations from the DOCX (round-trip through bytes).
//   AnnotationManager takes WmlDocument, so save the session first.
var saved = session.Save();
var wmlDoc = new WmlDocument("session.docx", saved);
var annotations = AnnotationManager.GetAnnotations(wmlDoc);
var target = annotations.First(a => a.LabelId == "INDEMNIFICATION");

// Step 2 — walk the bookmark to find the OOXML elements it covers,
//   then map their Unids back to AnchorTargets in the live projection.
using var probe = new DocxSession(saved);  // matches the saved state
var bookmarkName = target.BookmarkName;     // "_Docxodus_Ann_<id>"
var index = probe.Project().AnchorIndex;

// Find the bookmarkStart / bookmarkEnd in the body. (Pseudocode — the
// AnchorTarget.Resolve walk is private; copy the W.bookmarkStart match
// logic from AnnotationManager.GetAnnotationsInternal until the helper
// API lands.)
var anchors = index.Values
    .Where(t => /* element ancestry contains a w:bookmarkStart
                   with the matching name and a preceding w:bookmarkEnd
                   sibling within the body */)
    .ToList();

// Step 3 — hand anchors to the mutation surface.
foreach (var a in anchors)
    session.ReplaceText(a.Anchor.Id, "Revised indemnification language…");
```

Two caveats to internalize while [#132](https://github.com/JSv4/Docxodus/issues/132) is open:

- **`AnnotationManager` operates on `WmlDocument`, not the long-lived session.** The round-trip costs a save + reopen. Acceptable for low-frequency lookups; not acceptable to do per-mutation.
- **Bookmarks can span partial paragraphs.** A bookmark that starts mid-run will still return the enclosing block's anchor. That's usually what an agent wants ("edit this paragraph"), but for surgical character-range edits you'll need to combine `FindByAnnotation` with `ApplyFormat(anchor, CharSpan, op)` after computing the offset by walking the bookmark range yourself.

The agent's prompt should also be aware: it can call `AnnotationManager.GetAnnotations(doc)` once at the start of a session to enumerate available labels (e.g., "you can target: INDEMNIFICATION, TERMINATION, GOVERNING_LAW") and present those as tools rather than asking the LLM to discover them from text.

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
```

### Performance

~400 ms for a full-document grep over the 150 KB NVCA Model COI (~500 anchors, ~31 underscore-placeholder matches). Scales linearly with document size + match count.

### Known limits

- **Each block is grep'd in isolation.** Grep iterates paragraphs/headings/list-items and runs the regex against each one's flat text independently — there's no document-wide text stream that flows across paragraph boundaries. So `session.Grep("Hello world")` won't match if `"Hello "` is in one paragraph and `"world"` is in the next, even though they appear adjacent in the rendered doc. This is by design: every `TextMatch` carries a single `EnclosingAnchor` for the caller to hand back to `ReplaceText`/`Raw.ReplaceXml`, and matches that crossed block boundaries would have no single enclosing anchor. If you actually need cross-block search (legal clauses split for readability, multi-paragraph regions, etc.), the workaround until [#146](https://github.com/JSv4/Docxodus/issues/146) lands is to regex against the full markdown projection and walk `{#kind:scope:unid}` tokens back to anchors:
    ```csharp
    var proj = session.Project();
    foreach (Match m in Regex.Matches(proj.Markdown, multiLinePattern, RegexOptions.Singleline))
    {
        // m.Index / m.Length point into proj.Markdown; surrounding anchor tokens identify the blocks involved.
    }
    ```
- `RegexOptions` is the .NET enum; the npm wrapper passes its numeric value through (see `GrepOptions` in `npm/src/types.ts`).
- Tracked-change content currently follows the projector's accepted/rendered text — `Settings.TrackedChanges = StripDeletions` won't filter `<w:del>` content out of Grep yet. Worth opening as a follow-up if it matters.

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

## Error catalog (by remediation)

Errors are grouped by what the agent should do in response, not by where in the code they're raised. The `EditErrorCode` enum lives in `Docxodus/DocxSession.cs`; the snake-case TypeScript union is in `npm/src/types.ts`.

| The agent should… | When it sees these codes |
|---|---|
| Re-project and re-derive the anchor from current text | `AnchorNotFound` |
| Re-read the anchor's kind via `GetAnchorInfo`, reissue with the right op or coordinates | `AnchorWrongKind`, `AnchorsNotAdjacent`, `InvalidPosition`, `OffsetOutOfRange` |
| Fix the markdown payload (the message names what's wrong) | `MalformedMarkdown`, `UnsupportedMarkdownSyntax`, `AnchorTokenInPayload` |
| Call the v1 op the message names, or fall back to `Raw.InsertXml` | `TableInsertNotSupported`, `FootnoteRefNotSupported`, `CommentMarkerNotSupported`, `ImageInsertNotSupported` |
| Re-query (no `ListStyles()` API in v1; the agent guesses from the projection) | `UnknownStyle`, `InvalidListLevel` |
| Use `Raw.GetXml(anchor)` as a template, mutate, resubmit | `MalformedXml`, `DisallowedNamespace`, `IncompatibleElementType`, `ValidationFailed` |
| Stop, reopen, or accept "no more history" | `SessionDisposed`, `NothingToUndo`, `NothingToRedo` |
| Should not happen; treat as a bug. Op is rolled back, safe to retry once or report. Full exception is on `session.LastInternalError` | `InternalError` |

**Failure is transactional.** On any error, no mutation was applied. The pre-op snapshot was taken but is discarded without restoring (because nothing landed in the first place). Failed ops do not consume an undo slot. This holds for both pre-apply validation failures and runtime failures caught and rolled back.

## Recipes

These are worked examples drawn from the end-to-end smoke test (`DocxSessionSmokeTest.cs::DS999`) and the per-tier tests, lightly genericized. They use the .NET API; the TypeScript API is shape-identical (camelCase method names, `string` anchors, `Promise`-free synchronous returns from the npm wrapper since everything runs on the WASM worker).

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

## Known limits and open questions

- **`MarkdownPatch.Markdown` is currently the full re-projection.** The `ScopeAnchorId` field correctly identifies the smallest enclosing block, but the payload is the whole document re-projected. A future optimization (per the spec's open questions) is to emit only the markdown for the named scope. Cheap mitigation: callers that care can splice using their cached projection.
- **Snapshot granularity is per-part XML clone.** For documents with very large embedded images or huge tables, per-element diffs would be more memory-efficient. Deferred until measured to be a problem.
- **No `ListStyles()` query API in v1.** Agents must guess `styleId` values for `SetParagraphStyle` from what they see in the projection. `Heading1`–`Heading6`, `Quote`, and `Code` are reliable defaults across most documents.
- **Closing a session mid-flight from JS.** The WASM bridge holds sessions in a static dictionary keyed by handle; if a JS caller drops a `DocxSession` without calling `close()`, the .NET-side session is not eligible for GC. The npm wrapper exposes `Symbol.dispose` for TypeScript 5.2+ `using` blocks; older runtimes need explicit `.close()`.
- **`Save()` strips internal `PtOpenXml:Unid` attributes by default.** The projector assigns a Unid to every descendant of every projected scope; persisting them grows large documents by hundreds of KB of attribute noise (a 148 KB NVCA Model COI round-tripped at 588 KB before this default flipped). Anchor ids therefore do **not** survive `Save` → re-open by default — a fresh session re-assigns Unids and gets new ids. Set `DocxSessionSettings.PersistAnchorIds = true` to keep the ids (which keeps the bloat). This resolves Open Question #1 in `markdown_projection.md` in favor of "clean OOXML out by default, opt in to anchor stability."

## Related

- [`markdown_projection.md`](markdown_projection.md) — the read-side projector this builds on (anchor scheme, scope semantics, projector handlers)
- [`docx_converter.md`](docx_converter.md) — `WmlToHtmlConverter` internals (sibling write-side converter with very different goals)
- [`tracked_changes.md`](tracked_changes.md) — informs the `TrackedChangeMode` setting
- [`incremental_annotation_overlay.md`](incremental_annotation_overlay.md) — anchor-based overlay pattern; the read-side analog of this write-side API
