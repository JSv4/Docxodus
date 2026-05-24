# Markdown Projection

> **Status:** Implemented. All eight phases of the implementation plan landed; tests under `Docxodus.Tests/WmlToMarkdownConverterTests.cs` cover the surface and the WASM/npm bridge ships `convertWmlToMarkdown`. The "Implementation Plan (Phases)" section below documents the staging order that produced today's code.

The markdown projection is a deterministic, **anchor-addressed** rendering of a DOCX as Markdown. It is a sibling to `WmlToHtmlConverter` and `OpenContractExporter` in the converter family, intended as a substrate for tooling that wants to operate on Word documents the way it would operate on source files — search, splice, diff, address by ID. Use cases include LLM-driven editing pipelines, structured search indexers, and diff/review UIs that need a text view richer than `WmlToHtmlConverter` strips down to.

## Goals

1. **Stable addressing.** Every paragraph, heading, list item, table, table cell, comment, footnote, and endnote in the projection is reachable by an anchor that survives reformatting and reordering.
2. **Deterministic output.** Two runs on the same input produce byte-identical output.
3. **Round-trippable references.** An anchor in the projection maps unambiguously back to an `XElement` (or set of elements) in the underlying OOXML. The projection itself is read-only — mutation lives in callers — but the anchor → element resolver is part of this module's API.
4. **Lossy by design, honestly.** Anything that can't fit in Markdown becomes either a structured opaque anchor (with metadata callers can fetch via the SDK) or a clearly-marked annotation. Silent loss is a bug.

## Non-Goals

- **Round-trip rendering** (Markdown → DOCX). That problem already has `HtmlToWmlConverter` and `DocumentAssembler`; this converter is one-way.
- **GFM-perfect tables.** Word tables (merged cells, nested tables, cell-level shading, vertical text, …) exceed GFM. We render what fits and surface the rest as opaque anchors.
- **Preserving every formatting nuance.** Bold/italic/code/links/headings/lists/quotes carry over. Font sizes, colors, character spacing, etc. do not — they're recoverable by anchor lookup if needed.

## Anchor Scheme

Anchors derive from the `Unid` system Docxodus already maintains on paragraphs and runs (see `AssignUnidToAllElements` / the legacy migration notes in `CLAUDE.md`). They are stable across edits unless the underlying element is removed.

**Format:** `{#kind:scope:unid}` where:

| Field | Values | Meaning |
|---|---|---|
| `kind` | `p`, `h`, `li`, `tbl`, `tr`, `tc`, `cmt`, `fn`, `en`, `img`, `drw`, `unk` | Element type |
| `scope` | `body`, `hdr1`…`hdrN`, `ftr1`…`ftrN`, `fn`, `en`, `cmt` | Which part of the package |
| `unid` | 8–16 hex chars | Stable element identifier |

Examples:

- `{#h:body:a1b2c3d4}` — heading in the main body
- `{#p:hdr1:9f8e7d6c}` — paragraph in the first header part
- `{#tc:body:1a2b3c4d}` — table cell in the body
- `{#cmt:cmt:00ff11ee}` — comment in the comments part

Anchors appear at the start of the line they refer to (block-level) or as inline `{#…}` markers (inline annotations like comments anchored to a span).

### Why prefix-with-`#`?

`{#…}` is the [Pandoc / kramdown attribute syntax](https://pandoc.org/MANUAL.html#extension-header_attributes). Most Markdown renderers either honor it (turning anchors into `id` attributes) or display it literally without breaking layout. Either way it survives copy/paste round trips through agent contexts.

## Element Coverage

| OOXML element | Markdown representation | Notes |
|---|---|---|
| `w:p` (no style) | Paragraph | Anchor on its own line above the text |
| `w:p` styled `Heading{1..9}` | `#`…`#########` heading | Heading level taken from style. ATX headings 7-9 exceed CommonMark; strict renderers degrade to literal text, but downstream parsers and LLMs recover the outline depth (silently clamping to `######` would lose it, which matters for legal/clause-numbered documents). |
| `w:p` styled `Heading*` + `w:numPr` | `#` heading with resolved number prefix | Auto-numbered headings (legal `FIRST: …` / `1.1 …` clause numbering) keep their resolved number; without this, headings render with only the trailing text. |
| `w:p` styled `Title` / `Subtitle` | `#` heading with class | |
| `w:p` with `w:numPr` | `-` or `1.` list item | Numbering resolved to literal markers; nested via indentation |
| `w:p` styled `Quote` / `IntenseQuote` | `>` blockquote | |
| `w:p` styled `Code` / `HTML Code` | Indented or fenced code block | |
| `w:r` with `w:b` | `**bold**` | |
| `w:r` with `w:i` | `*italic*` | |
| `w:r` with `w:rStyle="Code"` or monospace | `` `code` `` | |
| `w:r` with `w:strike` | `~~strike~~` | GFM extension |
| `w:hyperlink` | `[text](url)` | Internal links use anchor: `[text](#anchor)` |
| `w:tbl` (simple) | GFM pipe table | When no merged cells, nesting, or per-cell formatting |
| `w:tbl` (complex) | Opaque anchor block | `{#tbl:body:…}` followed by a fenced `text` block with a structural summary |
| `w:commentRangeStart`…`End` | Inline `{#cmt:cmt:…}` markers wrapping the commented span | Comment text appears in a Comments section at end |
| `w:footnoteReference` | `[^fn-xxxx]` GFM footnote ref | Definitions collected at end |
| `w:endnoteReference` | `[^en-xxxx]` | Same |
| `w:drawing` / `w:pict` (image) | `![alt](docxodus://img/…){#img:…}` | URL is a scheme the caller resolves; metadata accessible via anchor |
| `w:sdt` (content control) | Rendered content, anchor on outer SDT | The SDT itself is an anchor target so callers can address "this content control" |
| `w:ins` / `w:del` (tracked changes) | Configurable: accept, show as `{+ins+}`/`{-del-}`, or omit | Mirrors `WmlToHtmlConverter.RenderTrackedChanges` |
| `w:sectPr` | `---` thematic break with section anchor | Sections become navigable |

Anything not in the table above renders as a single line:

```
{#unk:body:…} [unsupported: w:smartTag]
```

## Multipart Namespacing

A DOCX has many "documents" — the body is one part, but headers, footers, footnotes, endnotes, and comments live in sibling parts. The projection emits them as named sections so a single Markdown stream covers the whole package:

```markdown
# Document

{#p:body:…} The Provider shall...

---

# Headers

## hdr1
{#p:hdr1:…} CONFIDENTIAL

# Footers

## ftr1
{#p:ftr1:…} Page {PAGE} of {NUMPAGES}

# Footnotes

[^fn-aaaa]: {#fn:fn:aaaa} See Section 4.2 for definitions.

# Comments

- {#cmt:cmt:bbbb} **Alice** (2026-05-23): Should this be capitalized?
```

Callers that only care about the body can pass `Scopes = ProjectionScopes.Body` to skip the rest.

A `---` thematic break separates adjacent non-empty scope sections. Header/footer scopes whose only content is whitespace are suppressed entirely — DOCX files commonly declare 6+ header/footer parts for first-page/even-page/default variants, and emitting empty `## hdrN` titles for the unused variants would pad the projection with noise.

## Numbering Resolution

Word list numbers are computed from `w:numPr` referencing `numbering.xml` (`w:abstractNum` / `w:lvlText`), not stored as text. The projection **resolves numbering** so the agent sees what a human reads:

```markdown
{#li:body:…} 1. First item
{#li:body:…} 2. Second item
{#li:body:…}   a. Nested item
{#li:body:…} 3. Third item
```

The original `numPr` is recoverable through the anchor — callers that want to edit the list's numbering format address the source, not the rendered number.

Legal numbering (`1.1`, `1.1.1`, …) and other multi-level formats render verbatim.

## Tables

GFM pipe tables when:

- No merged cells (`w:gridSpan`, `w:vMerge`)
- No nested tables
- No per-cell formatting beyond bold/italic in cell content
- ≤ ~80 chars per cell (configurable)

Otherwise, an opaque anchor block:

````markdown
{#tbl:body:t1}
```table
rows: 4
cols: 3
caption: Fee Schedule
notes: merged cells in row 1; nested table in (3,2)
```
````

Per-cell content is reachable via `{#tc:body:…}` anchors that the caller can fetch individually with the SDK. The opaque block keeps the projection readable; the anchors keep it addressable.

## Round-Trip: Anchor → XElement

The companion API on the converter:

```csharp
var projection = WmlToMarkdownConverter.Convert(wmlDoc, settings);
// projection.Markdown is the text
// projection.AnchorIndex is an IReadOnlyDictionary<string, AnchorTarget>

var target = projection.AnchorIndex["p:body:a1b2c3d4"];
// target.PartUri, target.ElementXPath, target.Unid
// target.Resolve(WordprocessingDocument) -> XElement
```

This is the contract that makes the projection useful for editing: callers receive the projection *and* a way to walk back to the source for any anchor.

## Settings (Planned)

```csharp
public class WmlToMarkdownConverterSettings
{
    // What parts of the package to include.
    public ProjectionScopes Scopes = ProjectionScopes.All;

    // Heading level offset (e.g., 1 means Word Heading1 -> Markdown ##).
    public int HeadingLevelOffset = 0;

    // Inline anchors? Block-level only? Or omit?
    public AnchorRenderMode AnchorMode = AnchorRenderMode.Block;

    // How to handle complex tables.
    public TableRenderMode TableMode = TableRenderMode.GfmWithOpaqueFallback;

    // Max characters before a simple table becomes opaque.
    public int TableInlineCellMax = 80;

    // Tracked changes: accept silently, render as {+/-}, or strip dels.
    public TrackedChangeMode TrackedChanges = TrackedChangeMode.Accept;

    // Resolve list numbering to literal markers (default true).
    public bool ResolveNumbering = true;

    // Custom image URI scheme. Default: "docxodus://img/{unid}"
    public Func<ImageInfo, string>? ImageUriBuilder;
}

[Flags]
public enum ProjectionScopes
{
    Body = 1, Headers = 2, Footers = 4, Footnotes = 8, Endnotes = 16, Comments = 32,
    All = Body | Headers | Footers | Footnotes | Endnotes | Comments
}

public enum AnchorRenderMode { Block, BlockAndInline, None }
public enum TableRenderMode { GfmWithOpaqueFallback, AlwaysGfm, AlwaysOpaque }
public enum TrackedChangeMode { Accept, RenderInline, StripDeletions }
```

## Implementation Plan (Phases)

1. **Anchor index.** Walk the document, assign/reuse Unids on every block-level element across all parts, build the `AnchorIndex`. No markdown output yet — just verify uniqueness and round-trip resolution.
2. **Plain paragraphs + headings.** Emit body paragraphs and styled headings with anchors. Tests against `TestFiles/HC*` documents.
3. **Inline runs.** Bold, italic, code, strike, hyperlinks.
4. **Lists with resolved numbering.** Lean on existing `ListItemRetrieverSettings` infrastructure.
5. **Simple tables (GFM) + opaque fallback.**
6. **Multipart parts.** Headers, footers, footnotes, endnotes, comments.
7. **Tracked changes rendering modes.**
8. **WASM + npm wrapper.** Add `[JSExport]` methods and TypeScript types matching the other converters.

Each phase ships with tests in `Docxodus.Tests/WmlToMarkdownConverterTests.cs` (test prefix `MD###`).

## Getting Started (For Implementers)

### Worked Example

Given a DOCX whose body contains, in order:

1. A paragraph styled `Heading1` with text "Indemnification"
2. A paragraph (Normal style) with text "The Provider shall indemnify..."
3. A two-item bulleted list with text "First obligation" / "Second obligation"

The default-settings projection should produce:

```markdown
# Document

{#h:body:a1b2c3d4} # Indemnification

{#p:body:e5f6a7b8} The Provider shall indemnify...

{#li:body:c9d0e1f2} - First obligation
{#li:body:3a4b5c6d} - Second obligation
```

Notes on the example:
- Anchor token sits on its own line preceding the block, separated by a single space from the markdown that follows.
- The `# Document` header that opens the body scope is fixed (see Multipart Namespacing); no anchor — it's a scope marker, not addressable content.
- Unid values are 32-char hex (shortened above for readability); they come from `Guid.NewGuid().ToString().Replace("-", "")`.
- A blank line separates blocks. Lists are not separated internally.

Phase 2's first test (`MD001_HeadingAndParagraph`) should round-trip a fixture matching this shape and assert the markdown is byte-identical to the expected string above (with the actual Unids substituted via a helper that reads them off the source DOCX after assignment).

### Phase 1 in Detail

**Goal:** every block-level element in every part of the package has a stable Unid, and `MarkdownProjection.AnchorIndex` lets you walk back from any anchor to its `XElement`.

**Reuse, don't reinvent.**

- Unid attribute name: `PtOpenXml.Unid` (defined in `Docxodus/PtOpenXmlUtil.cs`).
- The existing Unid-assignment logic lives in two places, neither of which is general-purpose:
  - `WmlComparer.AssignUnidToAllElements(XElement)` — private to `WmlComparer.cs:8655`. Walks descendants, adds a Unid where missing. Also handles the special `w:footnote`/`w:endnote` case where the container itself needs a Unid.
  - `WmlToXmlUtil.AssignUnidToBlc` — block-level only, on a `WmlDocument` or `WordprocessingDocument`.
- For Phase 1, **extract the `WmlComparer` helper into a shared internal utility** (e.g. `UnidHelper.AssignToAllBlockElements`) and call it from both `WmlComparer` and the new converter. Don't duplicate.

**First deliverable (mergeable on its own):**

1. Extract the Unid helper.
2. Implement `WmlToMarkdownConverter.Convert(...)` such that it returns a `MarkdownProjection` with `Markdown = ""` and a populated `AnchorIndex` for every block-level element (`w:p`, `w:tbl`, `w:tr`, `w:tc`) across body, headers, footers, footnotes, endnotes, and comments.
3. Implement `AnchorTarget.Resolve(WordprocessingDocument)` — given the anchor's `PartUri` and `Unid`, walk that part's XML to find the matching element.
4. Tests:
   - `MD001_AnchorIndexIsExhaustive` — every `w:p`/`w:tbl` in a fixture is reachable by some anchor.
   - `MD002_AnchorsAreStable` — projecting twice gives the same anchor ids.
   - `MD003_AnchorsResolve` — every anchor in the index round-trips through `Resolve` back to a non-null `XElement` whose `PtOpenXml.Unid` attribute matches.
   - `MD004_AnchorsSurviveRoundTrip` — load → project → save (via `OpenXmlMemoryStreamDocument`) → re-load → re-project produces the same anchor ids for unchanged elements. **This is the doc's open question #1; if it fails, the converter must persist Unids back to the document before returning.**

### Structural Template to Mirror

The new converter's file organization should follow `Docxodus/WmlToHtmlConverter.cs`:

- One public static class with `Convert` entry points.
- Per-element handler methods (`WmlToHtmlConverter.ProcessParagraph` at line 4279 is the model — see the dispatch in `ConvertToHtmlTransform` around line 2493).
- Recursive descent driven by element name; each handler returns the rendered fragment (string for markdown; for HTML it returns `object`/`XElement`).
- Settings passed through every level — don't capture in fields on the static class.

Don't copy `WmlToHtmlConverter`'s ~6000-line bulk; copy its *shape*. The markdown converter will be much smaller because it intentionally drops most styling.

### WASM / npm Propagation (Phase 8)

When Phase 7 is merged, follow the existing pattern from `OpenContractExporter`:

1. Add `[JSExport]`-decorated methods to `wasm/DocxodusWasm/DocumentConverter.cs` that take/return JSON-serializable shapes (you can't return `XElement` or `IReadOnlyDictionary` directly — flatten to `MarkdownProjectionDto`).
2. Add TypeScript types and a wrapper in `npm/src/types.ts` + `npm/src/index.ts`.
3. Update the `DocxodusWasmExports` interface.
4. Build with `cd npm && npm run build`, then add Playwright tests under `npm/tests/`.

CLAUDE.md's "Feature Development Workflow" section is the canonical checklist for the cross-layer ripple.

### Performance Budget (Targets, Not Hard Constraints)

- Anchor index for a 100-page DOCX: < 200ms cold.
- Full projection (Phase 2+): < 1s for a 100-page DOCX, < 5s for 500 pages.
- Memory: O(document size), no full DOM duplication; reuse the OOXML SDK's `WordprocessingDocument` walk.

If Phase 1 measurements are >2× these numbers, surface in the PR and discuss before adding more functionality.

## Open Questions

- **Anchor stability across re-serialization.** Today Unids are assigned when needed; we should verify they survive `OpenXmlMemoryStreamDocument` round trips and document the lifecycle. If they don't, the converter must persist them back to the document.
- **Comment-on-span granularity.** Word comments can anchor to a span that crosses runs and paragraphs. Inline `{#cmt:…}` markers handle intra-paragraph; cross-paragraph spans probably need a start/end pair.
- **Images.** The placeholder `docxodus://img/{unid}` URI scheme assumes the caller provides resolution. Should we instead emit data URIs? Configurable via `ImageUriBuilder`, default TBD.
- **Bidirectional editing.** This converter is one-way, but the eventual `MarkdownToWml` story (or, more likely, an `ApplyEdit(anchor, op)` API) needs to be designed alongside its callers — not in isolation here.

## Related

- [`docx_converter.md`](docx_converter.md) — `WmlToHtmlConverter` internals (the sibling converter this most resembles)
- [`opencontracts_export.md`](opencontracts_export.md) — Other "structured export" precedent
- [`incremental_annotation_overlay.md`](incremental_annotation_overlay.md) — Anchor-based overlay pattern used by `ExternalAnnotationProjector`
- [`tracked_changes.md`](tracked_changes.md) — How tracked changes inform the `TrackedChangeMode` setting
