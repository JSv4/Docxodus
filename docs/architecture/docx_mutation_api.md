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
| `DeleteBlock(p)` (or `h`/`li`/`tbl`) | — | `p` + all descendant anchors | — | nearest stable ancestor |
| `DeleteBlock(fn)` / `DeleteBlock(en)` / `DeleteBlock(cmt)` | — | the definition anchor (and any cross-references it pointed at — those become "gone" but aren't separately addressed) | — | nearest stable ancestor in the body |
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

Tracked-change mode (v1): `DeleteRange` does a structural delete regardless of
`Settings.TrackedChanges`. Wrapping every run across many blocks in `w:del` is
deferred until a consumer needs it.

### `DeleteSection` — heading-bounded bulk removal

`session.DeleteSection(headingAnchorId)` deletes a heading and every sibling
below it up to (but not including) the next heading at the same or higher
level. "Level" matches the projection's notion: `Heading1` = 1, `Heading2` = 2,
…, `Title` = 1, `Subtitle` = 2.

If the target heading has no sibling-heading boundary after it, the section
extends to the end of the parent.

Built on `DeleteRange` semantics: same undo, same EditResult shape, same v1
tracked-change limitation.

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
- **Body scope only.** Bookmarks in headers/footers/footnotes aren't part of the v1 surface — `AnnotationManager` only writes to the main document part today. If header/footer annotation support lands, the helpers will return those anchors too.

Two caveats that are explicitly out of scope for v1 (tracked in [#132](https://github.com/JSv4/Docxodus/issues/132)):

- **Bookmarks that span partial paragraphs return the enclosing block's anchor**, not a character span. A character-range surgical edit needs `ApplyFormat(anchor, CharSpan, op)` after computing the offset within the bookmark range yourself.
- **Mutations don't auto-update bookmarks.** A `ReplaceText` / `SplitParagraph` / `MergeParagraphs` call can invalidate the bookmark covering the affected region. Bookmark preservation across mutations is a separate follow-up.

The agent's prompt should also be aware: it can call `session.ListAnnotations()` once at session start to enumerate available labels (e.g., "you can target: INDEMNIFICATION, TERMINATION, GOVERNING_LAW") and present those as tools rather than asking the LLM to discover them from text.

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
// summary.Filled / .Skipped / .Passes / .Unfilled / .Errors
```

What `FillPlaceholders` does internally that the recipe doesn't:

- **Reverse-offset ordering.** Earlier-offset matches in the same paragraph go stale once a later edit lands; `FillPlaceholders` sorts every pass by `(anchorId desc, span.Start desc)` so each block's matches are applied right-to-left.
- **`$`-prefix preservation.** The placeholder regex `\$?\[…\]` captures `$[___]` including the leading `$`. With `FillOptions.PreserveDollarPrefix = true` (default), the picker's return value gets `$` prepended when needed so `"0.20"` lands as `$0.20`, not `0.20`.
- **Multi-pass iteration.** `FindPlaceholders` returns innermost brackets only; stripping the inner can surface a previously-nested outer. The loop re-finds placeholders each pass and stops when a pass makes zero changes (or `MaxPasses` — default 8 — is hit).

The picker is invoked for every kind in `FillOptions.Kinds`, which defaults to `PlaceholderKinds.All` — so a picker that wants to ignore alternative-clause brackets should return `null` for them rather than relying on the option to filter them out. Set `Kinds = BlankFill | Instruction` if you want the prior behavior of leaving alternative clauses untouched.

The picker is invoked once per placeholder per pass; return `null` to skip. `BulkEditResult.Unfilled` lists every placeholder the picker said `null` to (deduplicated across passes). `BulkEditResult.Passes` is the highest iteration pass that actually filled at least one placeholder (so a single-fill convergence reports `Passes = 1`, not 2).

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

`DiffFormat.Unified` and `DiffFormat.SideBySide` are reserved for v2 (line-based diff)
— they throw `NotSupportedException` in v1. See issue #178.

## ReplaceTextRange — surgical text edits

`session.ReplaceTextRange(anchorId, find, replace, options?)` finds every literal occurrence of `find` in one paragraph/heading/list-item's flat text and substitutes `replace` for each, returning an `EditResult` per attempted match. Built on `Grep` — same fragment walker, opposite direction.

Three entry points covering the natural workflows:

```
session.ReplaceTextRange(anchor, find, replace, opts?)         // most common: replace every match in one block
session.ReplaceMatch(textMatch, replace)                       // convenience for a Grep result
session.ReplaceTextAtSpan(anchor, spanStart, spanLength, repl) // exact-span variant when several identical needles share a block
```

`ReplaceOptions`: `IgnoreCase` (case-insensitive find) and `MaxReplacements` (cap on how many to apply).

### Formatting-preservation contract

The replacement text inherits the formatting of the FIRST run the match spanned. Middle and trailing runs keep their `w:rPr` but lose the slice of text the match consumed — so a bold phrase that got partially overwritten still has bold formatting for any surviving text. This is the practical sweet spot: it solves the template-fill case where you want `[___]` → `Bluth Co.` to take on the surrounding text's formatting, and it's predictable for cross-formatting matches.

If you need different per-fragment behavior (e.g., the replacement should be bold even when the first fragment was plain), use `Grep` + bespoke `Raw.GetXml` mutation today, or wait for a future inline-markdown-aware overload.

### Ordering and atomicity

Multiple matches in the same paragraph are applied in **reverse document order** so each earlier-offset match's span stays valid after later edits land — the same trick the projector uses for tracked-change accept passes. The whole call records **one** snapshot; `Undo()` rolls every replacement back together.

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

For batched lookups (an agent that just enumerated 50 anchors and wants
previews for all of them), use `session.GetAnchorInfos(ids)` — a single pass
over the AnchorIndex instead of one walk per id. Returns
`IReadOnlyDictionary<string, AnchorInfo?>` — unknown ids map to null.

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
