# Document IR — Detailed Specification

**Date:** 2026-06-11
**Status:** Draft for Phase 1 implementation.
**Companion plan:** [`2026-06-11-ir-diff-layout-program-plan.md`](./2026-06-11-ir-diff-layout-program-plan.md)

## 1. Purpose, consumers, non-goals

The Document IR is a typed, normalized, anchor-identified, **immutable**
in-memory model of a Word document, built once per open and consumed by every
module that today re-derives its own private view from raw OOXML.

**Planned consumers, in onboarding order:**

1. Markdown projection (Phase 1 gate — validates the IR against shipped output)
2. Diff engine (Phase 2 — the IR's most demanding consumer; hashing and
   identity are designed for it)
3. `OpenContractExporter` text extraction (stretch)
4. Layout engine (Phase 3, deferred)

**Non-goals (v1):**

- **Not a writer.** The IR is a read-only projection; mutation stays in
  `DocxSession`'s XML path. (An IR→OOXML emitter arrives with the Phase 2
  revision renderer, scoped to that renderer's needs.)
- **Not lossless.** Unmodeled content is preserved as `Opaque` nodes with
  provenance, not modeled. Attempting full OOXML coverage is the documented
  death of projects like this.
- **Not the diff's tokenization.** Word splitting, case folding, and
  separator policy are *comparison settings*, not document facts. The IR
  stores runs; the diff engine tokenizes at diff time.
- **Not a new file format.** No IR serialization is a compatibility surface;
  the diagnostic JSON (§9) is for tests and debugging only.
- **Not public API (yet).** Everything is `internal` to the `Docxodus`
  assembly until an out-of-assembly consumer exists (Phase 2 productization
  at the earliest).

## 2. Design principles

- **P1 — Identity first.** Every block-level node carries the same
  deterministic, content-derived Unid-based anchor the markdown projection
  and `DocxSession` already use. Two opens of the same bytes produce
  identical anchors.
- **P2 — Normalize once.** All equality-relevant cleanup happens at read
  time, governed by the numbered rules in §5. Equality and hashing are
  defined on IR nodes; consumers never re-implement them.
- **P3 — Lossy-tolerant.** Anything unmodeled becomes `Opaque` with a
  canonical hash and a provenance pointer — diffable ("same bytes /
  different bytes / moved") without being understood.
- **P4 — Immutable snapshots.** An `IrDocument` never changes after `Read`.
  Thread-safe sharing and "compare two IRs" come for free.
- **P5 — Provenance everywhere.** Every node points back to its source
  `XElement`, so any consumer can drop to raw OOXML (`session.Raw`
  philosophy).
- **P6 — Resolved formatting is a view, not a mutation.** Direct properties
  are stored; effective (cascade-resolved) properties are computed lazily and
  cached. Nothing does what `FormattingAssembler` does to the XML.

## 3. Project layout & code standards

```
Docxodus/Ir/
  IrAnchor.cs          // identity types
  IrHash.cs            // hash value type + hashing helpers
  IrDocument.cs        // root + scopes
  IrBlocks.cs          // IrParagraph, IrTable, IrRow, IrCell, IrSectionBreak, IrOpaqueBlock
  IrInlines.cs         // IrTextRun, IrBreak, IrTab, IrFieldRun, IrHyperlink, IrNoteRef, IrInlineImage, IrOpaqueInline
  IrFormats.cs         // IrRunFormat, IrParaFormat, IrSectionFormat, IrListInfo
  IrRegistries.cs      // IrStyleRegistry, IrNumberingRegistry, IrThemeFonts
  IrNotes.cs           // footnote/endnote/comment stores
  IrReader.cs          // OOXML → IR entry point
  IrReaderOptions.cs
  IrNormalizer.cs      // rules N1–N15
  IrHasher.cs          // §6
  IrDiagnosticJson.cs  // §9
```

- Namespace `Docxodus.Ir`. All files `#nullable enable`.
- All types `internal`. Tests reach them via `InternalsVisibleTo`
  (add `Docxodus.Tests` if not already present).
- No SkiaSharp references; must compile under `WASM_BUILD`.
- C# `record` / `readonly record struct` throughout; collections exposed as
  `IReadOnlyList<T>`. **`Source` provenance fields are excluded from record
  equality** (declared as plain properties with `init`, compared by the
  hashes instead).

## 4. Core concepts

### 4.1 Identity — `IrAnchor`

```csharp
internal enum IrAnchorKind { P, H, Li, Tbl, Tr, Tc, Cmt, Fn, En, Img, Drw, Sec, Unk }

internal readonly record struct IrAnchor(IrAnchorKind Kind, string Scope, string Unid)
{
    public override string ToString() => $"{KindToken()}:{Scope}:{Unid}"; // "p:body:a1b2c3d4"
}
```

- `Kind`, `Scope`, and `Unid` use **exactly** the markdown projection's
  grammar (`{#kind:scope:unid}`, doc: `docs/architecture/markdown_projection.md`).
  Kind resolution for paragraphs (`p` vs `h` vs `li`) follows the projection's
  rules (outline level / list membership) so the same element gets the same
  anchor string from both code paths. **M1.4 asserts string equality of
  anchors between the IR path and the shipped projection.**
- Unids come from the existing deterministic assignment pipeline
  (`AddUnidsToMarkupInContentParts` / the `DocxSession` open path) — the IR
  reader reuses that code, it does not reinvent it.
- **Anchored:** paragraphs, tables, rows, cells, section breaks, footnotes,
  endnotes, comments, images/drawings, opaque blocks.
- **Not anchored:** runs and other inlines. Run identity is unstable across
  edits by nature; inline positions are addressed as (block anchor, char
  span), matching `DocxSession.ApplyFormat`'s existing addressing.

### 4.2 Hashes — `IrHash`

```csharp
internal readonly record struct IrHash // 32 bytes, SHA-256
{
    private readonly ulong _a, _b, _c, _d;
    public static IrHash Compute(ReadOnlySpan<byte> data);   // SHA-256
    public string ToHex();                                    // lowercase, for diagnostics
}
```

SHA-256 via `System.Security.Cryptography` (no new dependencies; speed is
adequate — hashing is O(document) once per open). Every block node carries:

- `ContentHash` — text identity (what the reader reads). §6.1.
- `FormatFingerprint` — formatting identity. §6.2.

The pair is the diff engine's primary signal: equal/equal → unchanged;
equal/different → format-only change; different → content change.

### 4.3 Provenance — `Source`

Every node has `XElement? Source` pointing into the `XDocument`(s) the reader
parsed, plus the owning part URI on block nodes. The `IrDocument` **pins**
those `XDocument` instances (a `Sources` property holding part URI →
`XDocument`) so provenance pointers stay alive exactly as long as the
snapshot. Memory consequence: an IR snapshot costs roughly (XML DOM) + (IR
nodes); budget in §10.

### 4.4 Opacity

Any element the reader does not model becomes:

```csharp
internal sealed record IrOpaqueBlock(IrAnchor Anchor, XName ElementName, IrHash CanonicalHash) : IrBlock;
internal sealed record IrOpaqueInline(XName ElementName, IrHash CanonicalHash) : IrInline;
```

`CanonicalHash` is computed over the canonicalized source XML (§6.3), so
opaque content participates correctly in diffing (unchanged / changed /
moved) without being understood. Opacity is the **default behavior for the
unknown**, never an exception. Promotion of an opaque element to a typed node
is an additive, snapshot-visible change (§11).

## 5. Reader semantics & normalization rules

### 5.1 Entry point

```csharp
internal sealed class IrReaderOptions
{
    public RevisionView RevisionView { get; init; } = RevisionView.Accept;
    public IrScopes Scopes { get; init; } = IrScopes.All;   // Body | HeadersFooters | Notes | Comments
}

internal enum RevisionView { Accept, Reject, FailIfPresent }

internal static class IrReader
{
    public static IrDocument Read(WmlDocument doc, IrReaderOptions? options = null);
}
```

`RevisionView` (rule N13): the v1 IR models a **revision-free** view. The
reader applies `RevisionProcessor` accept/reject to a working copy before
building (the original bytes are untouched). `Accept` is the default,
matching what `WmlToHtmlConverter` does today. Modeling in-flight `w:ins`/
`w:del` in the IR (needed for tracked-changes HTML on the IR path) is
deferred to v2 — tracked separately, not smuggled into v1.

### 5.2 Normalization rules

Each rule gets one unit test, named `IrNorm_N{nn}_*`. This table is the
project's single definition of document equality.

| # | Rule |
|---|------|
| N1 | Strip all `rsid*` attributes (`w:rsidR`, `w:rsidRPr`, `w:rsidRDefault`, …). |
| N2 | Drop `w:proofErr`, `w:noProof`, spelling/grammar markers entirely. |
| N3 | Drop `w:bookmarkStart`/`w:bookmarkEnd` from the node stream. Recoverable via `Source` on the containing block. (`_GoBack` and friends are pure noise; real bookmark modeling is a v2 candidate.) |
| N4 | Drop `w:lastRenderedPageBreak` (layout cache, not content). |
| N5 | Coalesce adjacent `IrTextRun`s with equal `IrRunFormat` into one run. Applied after all other inline rules. |
| N6 | `w:tab` → `IrTab`; `w:br` → `IrBreak(Kind: Line\|Page\|Column)`. Never folded into text. |
| N7 | `w:noBreakHyphen` → text U+2011; `w:softHyphen` → text U+00AD. |
| N8 | `w:sym` with a mappable char → text; otherwise `IrOpaqueInline`. |
| N9 | Fields: `w:fldSimple` and complex `w:fldChar begin/separate/end` sequences both become `IrFieldRun(Instruction, CachedResult: IReadOnlyList<IrInline>)`. The *cached result* feeds `ContentHash` (it is what a reader sees); the instruction is available to consumers and is **not** hashed. |
| N10 | Drop empty runs (no text after N1–N9) and empty `w:rPr`-only artifacts. |
| N11 | Text is preserved exactly as written, honoring `xml:space`. No whitespace conflation, no NBSP folding — `WmlComparer.ConflateBreakingAndNonbreakingSpaces`-style policies are diff-time settings. |
| N12 | `w:sdt` (content controls) and `w:smartTag` are unwrapped to their content. The block-level SDT's anchor lands on the unwrapped content's outer block (matching the projection's "anchor on outer SDT" behavior). SDT metadata recoverable via `Source`. |
| N13 | Revisions resolved per `RevisionView` before any node construction (§5.1). |
| N14 | `w:hyperlink` → `IrHyperlink(Target, Inlines)`; internal links carry the target anchor where resolvable. |
| N15 | Comment plumbing (`w:commentRangeStart/End`, `w:commentReference`) is removed from the inline stream and recorded in the comments store (§7.3) as (block anchor, char span) ranges. Comments never affect `ContentHash`. |

Theme/style indirection (e.g. `w:rFonts w:asciiTheme="minorHAnsi"`) is
resolved through `IrThemeFonts` into concrete font names **in the effective
format only**; direct-format records keep what the XML said.

## 6. Hashing specification

### 6.1 `ContentHash`

Computed over a canonical UTF-8 byte stream per block:

- `IrTextRun` → its text bytes.
- Non-text inlines → a single sentinel byte sequence `0x01 <kind-byte>`
  (`IrTab`=0x01, line/page/column break=0x02/0x03/0x04, note ref=0x05/0x06,
  image=0x07 followed by the image part's content hash, opaque=0x0F followed
  by its canonical hash). Sentinels are outside the Unicode text range so no
  text can collide with structure.
- `IrFieldRun` → the byte stream of its cached-result inlines (recursive),
  unbracketed: a field whose cached result reads "5" is content-equal to a
  literal "5" (deliberate — the hash captures what a reader sees; the
  instruction is consumer-visible but unhashed).
- `IrHyperlink` → sentinel `0x08`, the target string's UTF-8 bytes, sentinel
  `0x09`, then the child inlines' bytes. Linked text is therefore never
  content-equal to identical plain text, and a target change is a content
  change.
- `IrNoteRef` → its kind sentinel only (`0x05`/`0x06`), **without** the note
  id — ids are positional bookkeeping; note *content* equality is judged in
  the notes scope, and renumbering alone must not flip body hashes.
- Table: row sentinel `0x02 0x10`, cell sentinel `0x02 0x11`, then each
  cell's child-block content hashes in order; the table's `ContentHash` is
  the hash of that rollup. Paragraph hashes never leak across block
  boundaries.

Properties the diff engine relies on (asserted by tests): equal visible text
+ equal inline structure ⇔ equal `ContentHash`; formatting never affects it.

### 6.2 `FormatFingerprint`

Hash of the canonical serialization of the node's **direct** format record
(field name + value pairs in declaration order, omitted when null) plus the
`UnmodeledDigest` (§6.4). Run-level fingerprints roll up into the block
fingerprint together with the paragraph's own `IrParaFormat`, so "same text,
somebody bolded a word" flips the block fingerprint.

Deliberately **direct, not effective**: a style-definition edit changes every
paragraph's rendered appearance but should read as a style change, not N
paragraph edits. (The diff engine compares style definitions separately;
effective formats exist for consumers like the projection and future layout.)

### 6.3 Opaque canonicalization

For `Opaque` hashing: serialize the source element with attributes sorted by
(namespace, local name); strip `pt14`/PowerTools bookkeeping attributes and
rule-N1/N2 noise; normalize inter-element whitespace; UTF-8 encode; SHA-256.
This makes the hash stable across attribute reordering and rsid churn inside
content we don't model.

### 6.4 `UnmodeledDigest`

When the reader maps `w:rPr`/`w:pPr` into format records, any child element
it does **not** model is canonicalized (§6.3) into the record's
`UnmodeledDigest` field. Consequence: a format change in an unmodeled
property still flips the fingerprint — the diff reports "formatting changed"
without knowing *what* changed, instead of silently calling it equal. This is
the lossy-tolerance principle applied to formatting.

## 7. Type model

### 7.1 Document and scopes

```csharp
internal sealed record IrDocument
{
    public required IrScope Body { get; init; }
    public IReadOnlyList<IrHeaderFooter> Headers { get; init; }   // per-part, with type (default/first/even) + section linkage
    public IReadOnlyList<IrHeaderFooter> Footers { get; init; }
    public required IrNoteStore Footnotes { get; init; }          // note id → IrScope
    public required IrNoteStore Endnotes { get; init; }
    public required IrCommentStore Comments { get; init; }
    public required IrStyleRegistry Styles { get; init; }
    public required IrNumberingRegistry Numbering { get; init; }
    public required IrThemeFonts ThemeFonts { get; init; }
    public required IReadOnlyDictionary<Uri, XDocument> Sources { get; init; } // provenance pin

    public IrBlock? FindByAnchor(IrAnchor anchor);                // O(1), index built at read time
}

internal sealed record IrScope(string Name, IReadOnlyList<IrBlock> Blocks); // "body", "hdr1", "ftr1", "fn", "en", "cmt"
```

Scope names match the projection's multipart namespacing (`body`, `hdr1`,
`ftr1`, footnote/endnote/comment scopes) so anchors agree.

### 7.2 Blocks

```csharp
internal abstract record IrBlock
{
    public required IrAnchor Anchor { get; init; }
    public required IrHash ContentHash { get; init; }
    public required IrHash FormatFingerprint { get; init; }
    public XElement? Source { get; init; }                        // excluded from equality
}

internal sealed record IrParagraph : IrBlock
{
    public required IrParaFormat Format { get; init; }            // direct
    public IrParaFormat EffectiveFormat { get; }                  // lazy, cascade-resolved, cached
    public IrListInfo? List { get; init; }
    public required IReadOnlyList<IrInline> Inlines { get; init; }
}

internal sealed record IrTable : IrBlock
{
    public required IReadOnlyList<IrRow> Rows { get; init; }
    public required IrHash UnmodeledTablePropsDigest { get; init; } // tblPr/tblGrid via §6.3
}

internal sealed record IrRow(IrAnchor Anchor, IReadOnlyList<IrCell> Cells, IrHash ContentHash);
internal sealed record IrCell(IrAnchor Anchor, IReadOnlyList<IrBlock> Blocks,
                              int GridSpan, IrVMerge VMerge, IrHash ContentHash);

internal sealed record IrSectionBreak : IrBlock                   // kind = Sec
{
    public required IrSectionFormat Format { get; init; }         // page size, margins, orientation, type, hf refs
}
```

Cells contain blocks recursively — nested tables come for free. The body's
trailing `sectPr` becomes a final `IrSectionBreak`, so section structure is
uniform.

### 7.3 Inlines

```csharp
internal abstract record IrInline;

internal sealed record IrTextRun(string Text, IrRunFormat Format) : IrInline;
internal sealed record IrTab(IrRunFormat Format) : IrInline;
internal sealed record IrBreak(IrBreakKind Kind) : IrInline;              // Line | Page | Column
internal sealed record IrHyperlink(string? Target, IrAnchor? InternalTarget,
                                   IReadOnlyList<IrInline> Inlines) : IrInline;
internal sealed record IrFieldRun(string Instruction,
                                  IReadOnlyList<IrInline> CachedResult) : IrInline;
internal sealed record IrNoteRef(IrNoteKind Kind, string NoteId) : IrInline; // Footnote | Endnote
internal sealed record IrInlineImage(Uri PartUri, IrHash ImageBytesHash,
                                     long WidthEmu, long HeightEmu,
                                     string? AltText) : IrInline;
// + IrOpaqueInline (§4.4)
```

Comment stores:

```csharp
internal sealed record IrCommentStore(IReadOnlyList<IrComment> Comments);
internal sealed record IrComment(IrAnchor Anchor, string Author, string? Initials, string? Date,
                                 IReadOnlyList<IrBlock> Blocks,
                                 IReadOnlyList<IrCommentTarget> Targets);
internal sealed record IrCommentTarget(IrAnchor BlockAnchor, int StartChar, int EndChar); // N15
```

### 7.4 Formats

Modeled fields are the deliberate v1 subset — everything observable in the
markdown projection plus what the diff needs for `w:rPrChange`-grade
reporting. Everything else: `UnmodeledDigest`.

```csharp
internal sealed record IrRunFormat
{
    public string? StyleId { get; init; }
    public bool? Bold { get; init; }
    public bool? Italic { get; init; }
    public IrUnderline? Underline { get; init; }       // kind + color
    public bool? Strike { get; init; }
    public bool? DoubleStrike { get; init; }
    public IrVertAlign? VertAlign { get; init; }       // Subscript | Superscript
    public string? FontAscii { get; init; }            // as written; theme-resolved only in effective
    public int? SizeHalfPoints { get; init; }
    public string? ColorHex { get; init; }             // as written, including the literal "auto"
    public string? Highlight { get; init; }
    public bool? Caps { get; init; }
    public bool? SmallCaps { get; init; }
    public bool? Vanish { get; init; }
    public required IrHash UnmodeledDigest { get; init; }
}

internal sealed record IrParaFormat
{
    public string? StyleId { get; init; }
    public IrJustification? Justification { get; init; }
    public int? IndentLeftTwips { get; init; }
    public int? IndentRightTwips { get; init; }
    public int? IndentFirstLineTwips { get; init; }    // negative = hanging
    public int? SpacingBeforeTwips { get; init; }
    public int? SpacingAfterTwips { get; init; }
    public IrLineSpacing? LineSpacing { get; init; }   // value + rule
    public int? OutlineLevel { get; init; }
    public bool? KeepNext { get; init; }
    public bool? KeepLines { get; init; }
    public bool? PageBreakBefore { get; init; }
    public required IrHash UnmodeledDigest { get; init; }
}

internal sealed record IrListInfo(int NumId, int? AbstractNumId, int Ilvl,
                                  string NumberFormat, int? StartOverride, bool FromStyle);
// AbstractNumId is null until numbering resolution lands (M1.3) — null means "not yet
// resolved", never a sentinel that could collide with a real abstractNumId.
```

`IrListInfo` carries exactly the facts `GetBlockMetadata`/`GetListMembership`
report today (`numId`/`abstractNumId`/`ilvl`/format/start-override/
from-style); M1.3 asserts parity with that surface.

## 8. Immutability, laziness, threading

- All construction happens inside `IrReader.Read`; nothing mutates after.
- The only lazy members are effective-format resolutions, implemented with
  `Lazy<T>`-style thread-safe caching keyed off the immutable registries —
  semantically pure, so laziness is unobservable.
- A snapshot may be shared freely across threads. Two snapshots of the same
  bytes are `Equals`-equal node-for-node (provenance excluded), and produce
  identical hashes and anchors — this **determinism guarantee is a tested
  invariant**, not an aspiration.

## 9. Diagnostic JSON

`IrDiagnosticJson.Write(IrDocument)` emits a stable, human-readable dump:
anchors, hashes (hex), format records, text, opaque element names. Used for:

- Golden-snapshot conformance tests over `TestFiles/` (M1.1 onward).
- Debugging ("show me what the IR thinks this document is").

It is explicitly **not** a versioned format; snapshots are regenerated under
review when the IR evolves (never blind-regenerated — every snapshot diff is
triaged, per the program plan's normalization-churn mitigation).

## 10. Conformance & performance budgets

| Check | Mechanism | Gate |
|---|---|---|
| Reader totality | `Read` over every `TestFiles/` fixture, no throws | M1.1 |
| Normalization rules | one unit test per N-rule | M1.2 |
| Hash stability/sensitivity | re-read equality; targeted mutations | M1.2 |
| Determinism | two reads of same bytes → node-equal, hash-equal, anchor-equal | M1.2 |
| List/metadata parity | vs `GetListMembership`/`GetBlockMetadata` | M1.3 |
| Anchor parity | string-equal vs shipped markdown projection | M1.4 |
| Projection equivalence | corpus diff vs shipped converter, triaged | M1.4 (phase gate) |
| Performance | IR build + projection ≤ 2× current converter wall time, corpus-wide; memory ≤ 3× document XML size | M1.4 (phase gate) |

## 11. Evolution policy

- **Opaque promotion:** modeling a previously-opaque element is additive but
  changes hashes and snapshots for affected documents; it lands with
  regenerated-and-reviewed snapshots and a CHANGELOG entry. Consumers must
  never persist IR hashes across library versions (they are session-scoped
  identities, like Unids).
- **v2 candidates, explicitly deferred:** revision-aware IR (`w:ins`/`w:del`
  as nodes, for tracked-changes HTML), bookmarks as ranges, content-control
  metadata as typed facts, textbox/shape body content, IR→OOXML writer
  beyond the Phase 2 renderer's needs.
- **Visibility:** types go `public` only when Phase 2 productization needs
  them, and then under a documented "experimental" banner.

## 12. Open questions (to resolve during M1.1–M1.2, not blockers)

1. Whether `IrHyperlink` should be a block-spanning wrapper or flattened with
   per-run link targets (current lean: wrapper, matching OOXML nesting).
2. ~~Comment target spans when the range crosses block boundaries — one
   `IrCommentTarget` per touched block vs a (startAnchor, endAnchor) pair.~~
   **Resolved (M1.3):** one `IrCommentTarget` per touched block. A cross-block
   range closes at the first block's end offset and re-opens at offset 0 of each
   subsequent block until the `commentRangeEnd`, so a range spanning N blocks
   yields N targets. Char offsets count visible `IrTextRun` characters only
   (tabs/breaks/images/fields/opaque inlines count 0) — stable under the N5
   coalescing pass. Orphan range-starts are discarded; a `commentReference` for a
   comment with no ranges records a zero-length target at the reference offset.
3. Whether `IrSectionBreak` should also surface header/footer *content*
   linkage or just references (lean: references in v1; layout needs content
   linkage but layout is deferred).
4. Image identity: `ImageBytesHash` of the part vs relationship id — bytes
   hash chosen so the diff sees "same image re-added" as equal; confirm cost
   is acceptable on image-heavy fixtures.
