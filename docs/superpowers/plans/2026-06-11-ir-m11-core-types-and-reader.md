# Document IR — M1.1 Core Types + Reader Skeleton Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Implement Phase 1 / M1.1 of the Document IR: value types, node model, hashing, an `IrReader` covering body paragraphs/runs/tables/breaks/tabs with opaque fallback, diagnostic JSON, and corpus totality + golden snapshot tests.

**Architecture:** Per `docs/superpowers/specs/2026-06-11-document-ir-spec.md` (the spec — authoritative). New `internal` namespace `Docxodus.Ir` inside the existing `Docxodus` project. Immutable records; provenance excluded from equality; everything unmodeled becomes `Opaque` nodes with canonical hashes. Reuses the existing deterministic Unid pipeline (`UnidHelper.AssignToAllElementsDeterministic`) and the markdown projection's kind resolution (`KindFor`, to be made `internal`).

**Tech stack:** .NET 8, xUnit, System.Text.Json (in-box), `System.Security.Cryptography.SHA256`. No new package references. All files `#nullable enable`. Must compile under `WASM_BUILD` (no SkiaSharp).

**M1.1 scope reductions (deliberate, per program plan — do NOT implement these now):**
- Scopes: **body only**. Headers/footers/footnotes/endnotes/comments stores exist as empty placeholder instances (M1.3).
- Registries (`IrStyleRegistry`, `IrNumberingRegistry`, `IrThemeFonts`): empty placeholder records; no effective-format resolution (M1.3). `IrParagraph.EffectiveFormat` is **omitted entirely** in M1.1.
- Normalization: only N1 (via canonicalization), N5 (run coalescing), N6 (tab/br), N10 (empty runs), N11 (exact text), N13 (RevisionView). Fields, hyperlinks, sym, SDT, comment plumbing (N7–N9, N12, N14, N15) become `IrOpaqueInline` for now (M1.2).
- `IrListInfo` is constructed with raw `numPr` facts only where trivially readable (`numId`, `ilvl` from direct pPr); `abstractNumId`/format/start-override/from-style are M1.3 — use `-1`/`""`/`null`/`false` placeholders and document it.
- Mid-document section breaks (sectPr inside pPr) stay inside the paragraph's `UnmodeledDigest`; only the body-trailing `sectPr` becomes an `IrSectionBreak`.

**Verification commands** (used by every task):

```bash
dotnet build Docxodus.sln                                   # must be clean
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~Docxodus.Tests.Ir"
```

Run the full suite (`dotnet test Docxodus.Tests/Docxodus.Tests.csproj`) before each commit to prove no regression.

---

## Task 1: Value types — `IrHash`, `IrAnchor`, `IrProvenance`

**Files:**
- Create: `Docxodus/Ir/IrHash.cs`
- Create: `Docxodus/Ir/IrAnchor.cs`
- Create: `Docxodus/Ir/IrProvenance.cs`
- Test: `Docxodus.Tests/Ir/IrValueTypeTests.cs`

**`IrHash`** — `internal readonly struct`, 32 bytes stored as four `ulong` fields. Members:
- `static IrHash Compute(ReadOnlySpan<byte> data)` — SHA-256 via `System.Security.Cryptography.SHA256.HashData`.
- `static IrHash Compute(string text)` — UTF-8 convenience overload.
- `string ToHex()` — 64-char lowercase hex.
- Full value equality (`IEquatable<IrHash>`, `==`/`!=`, `GetHashCode`).

**`IrAnchor`** — per spec §4.1:

```csharp
internal enum IrAnchorKind { P, H, Li, Tbl, Tr, Tc, Cmt, Fn, En, Img, Drw, Sec, Unk }

internal readonly record struct IrAnchor(IrAnchorKind Kind, string Scope, string Unid)
{
    public override string ToString() => $"{KindToken(Kind)}:{Scope}:{Unid}";
    public static string KindToken(IrAnchorKind kind);            // P→"p", H→"h", Li→"li", Tbl→"tbl", Tr→"tr", Tc→"tc", Cmt→"cmt", Fn→"fn", En→"en", Img→"img", Drw→"drw", Sec→"sec", Unk→"unk"
    public static IrAnchorKind KindFromToken(string token);       // inverse; throws ArgumentException on unknown token
}
```

Token strings MUST match `WmlToMarkdownConverter.KindFor`'s vocabulary exactly (see `Docxodus/WmlToMarkdownConverter.cs:537`); anchor string format matches the projection's `Anchor.Id` (`kind:scope:unid`, unid = 32-char hex).

**`IrProvenance`** — the provenance-excluded-from-equality mechanism (spec §3): a small `internal sealed class` holding `XElement? Element { get; init; }` whose `Equals(object?)` returns `true` for any other `IrProvenance` and `GetHashCode()` returns `0`. Records that embed an `IrProvenance Source` property therefore compare equal regardless of provenance. XML-doc this trick explicitly.

**Tests (write first, watch fail, then implement):**
- `IrHash_Compute_MatchesKnownSha256Vector` — `Compute("abc")` hex equals `ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad`.
- `IrHash_Equality_ByValue` — two computes of same bytes equal; different bytes not equal.
- `IrAnchor_ToString_MatchesProjectionGrammar` — `new IrAnchor(IrAnchorKind.P, "body", "a1b2…")` renders `p:body:a1b2…`.
- `IrAnchor_KindTokens_RoundTrip` — every enum value round-trips through token and back.
- `IrProvenance_NeverAffectsEquality` — two instances wrapping different `XElement`s are `Equals`-equal.

**Steps:** failing tests → run (`--filter "FullyQualifiedName~IrValueTypeTests"`, expect compile failure/red) → implement → green → full suite → commit `feat(ir): add IrHash, IrAnchor, IrProvenance value types`.

---

## Task 2: Format records and enums

**Files:**
- Create: `Docxodus/Ir/IrFormats.cs`
- Test: `Docxodus.Tests/Ir/IrFormatTests.cs`

Implement exactly the spec §7.4 records plus supporting enums — `IrRunFormat`, `IrParaFormat`, `IrListInfo`, and:

```csharp
internal enum IrUnderlineKind { Single, Double, Thick, Dotted, Dashed, Wave, Words, None, Other }
internal sealed record IrUnderline(IrUnderlineKind Kind, string? ColorHex);
internal enum IrVertAlign { Subscript, Superscript }
internal enum IrJustification { Left, Center, Right, Both, Distribute, Other }
internal enum IrLineSpacingRule { Auto, AtLeast, Exact }
internal sealed record IrLineSpacing(int ValueTwips, IrLineSpacingRule Rule);
internal enum IrBreakKind { Line, Page, Column }
internal enum IrNoteKind { Footnote, Endnote }
internal enum IrVMerge { None, Restart, Continue }

internal sealed record IrSectionFormat
{
    public int? PageWidthTwips { get; init; }
    public int? PageHeightTwips { get; init; }
    public bool? Landscape { get; init; }
    public int? MarginTopTwips { get; init; }
    public int? MarginBottomTwips { get; init; }
    public int? MarginLeftTwips { get; init; }
    public int? MarginRightTwips { get; init; }
    public string? SectionType { get; init; }       // w:type/@w:val as written
    public required IrHash UnmodeledDigest { get; init; }
}
```

`IrRunFormat`/`IrParaFormat`/`IrListInfo`: copy field-for-field from spec §7.4 (including `required IrHash UnmodeledDigest`). Unknown `w:u/@w:val` values map to `IrUnderlineKind.Other` (the raw value still lands in `UnmodeledDigest` — that wiring is Task 5's job, not this task's).

**Tests:**
- `IrRunFormat_RecordEquality_IncludesUnmodeledDigest` — identical fields + same digest → equal; same fields + different digest → not equal.
- `IrParaFormat_RecordEquality` — same shape check.
- `IrListInfo_Equality` — positional record equality sanity.

**Steps:** failing tests → red → implement → green → full suite → commit `feat(ir): add IR format records and enums`.

---

## Task 3: Node model — blocks, inlines, document, scopes, stores

**Files:**
- Create: `Docxodus/Ir/IrBlocks.cs`
- Create: `Docxodus/Ir/IrInlines.cs`
- Create: `Docxodus/Ir/IrDocument.cs`
- Create: `Docxodus/Ir/IrNotes.cs`
- Create: `Docxodus/Ir/IrRegistries.cs`
- Test: `Docxodus.Tests/Ir/IrNodeTests.cs`

Implement spec §7.1–§7.3 with these M1.1 adjustments:
- Every node's provenance is `public IrProvenance Source { get; init; } = new();` (NOT a raw `XElement` — Task 1's equality-neutral wrapper). Block-level nodes also get `public Uri? PartUri { get; init; }` inside the provenance wrapper — add `Uri? PartUri { get; init; }` to `IrProvenance` instead of the block, keeping all provenance equality-neutral.
- `IrParagraph` has NO `EffectiveFormat` member (M1.3).
- `IrRegistries.cs` contains empty placeholders: `internal sealed record IrStyleRegistry { public static readonly IrStyleRegistry Empty = new(); }` and likewise `IrNumberingRegistry`, `IrThemeFonts`.
- `IrNotes.cs`: `IrNoteStore` (empty-capable: `public static readonly IrNoteStore Empty`), `IrCommentStore.Empty`, plus `IrComment`/`IrCommentTarget` records per spec (constructed only in M1.3, but the types exist now so `IrDocument` is complete).
- `IrDocument`: per spec §7.1, with `Headers`/`Footers` as `IReadOnlyList<IrHeaderFooter>` where `internal sealed record IrHeaderFooter(string ScopeName, IrHeaderFooterKind Kind, IrScope Scope);` and `internal enum IrHeaderFooterKind { Default, First, Even }` — empty lists in M1.1. `FindByAnchor` backed by `required IReadOnlyDictionary<string, IrBlock> AnchorIndex` (keys are `IrAnchor.ToString()` strings), populated by the reader; `FindByAnchor(IrAnchor a)` does a dictionary lookup on `a.ToString()`.
- `Sources` property: `required IReadOnlyDictionary<Uri, XDocument>` per spec.

Inlines per spec §7.3 — note `IrTextRun`, `IrTab` carry `IrRunFormat Format`; `IrBreak` carries only `Kind`; include `IrHyperlink`, `IrFieldRun`, `IrNoteRef`, `IrInlineImage`, `IrOpaqueInline` types now (reader emits only TextRun/Tab/Break/OpaqueInline in M1.1, but the model is complete).

**Tests:**
- `IrParagraph_Equality_IgnoresProvenance` — two structurally identical paragraphs with different `Source` elements are equal.
- `IrTable_NestedCells_Construct` — build a 1×1 table whose cell contains a paragraph; assert shape.
- `IrDocument_FindByAnchor_ReturnsBlock` — hand-built document with one paragraph; lookup by anchor returns it; unknown anchor returns null.

**Steps:** failing tests → red → implement → green → full suite → commit `feat(ir): add IR node model (blocks, inlines, document, scopes)`.

---

## Task 4: Canonicalization and hashing — `IrHasher`

**Files:**
- Create: `Docxodus/Ir/IrHasher.cs`
- Test: `Docxodus.Tests/Ir/IrHasherTests.cs`

**`Canonicalize(XElement) → byte[]`** (spec §6.3): clone the element; recursively (a) drop attributes whose name is in the noise set — any attribute whose local name starts with `rsid`, anything in the `PtOpenXml` (pt14) namespace, `xmlns` declarations; (b) drop `w:proofErr`/`w:noProof` elements; (c) sort remaining attributes by (namespace, local name); (d) serialize with `SaveOptions.DisableFormatting`; UTF-8 encode. `CanonicalHash(XElement)` = `IrHash.Compute(Canonicalize(el))`.

**Content-hash stream builder** (spec §6.1): `internal sealed class IrContentHashBuilder` with append methods used by the reader:
- `AppendText(string)` — UTF-8 bytes. (No escaping needed: sentinel lead byte `0x01` cannot occur in XML text — XML 1.0 forbids U+0001; assert/document this.)
- `AppendSentinel(byte kind)` — writes `0x01, kind`. Kind bytes per spec §6.1: tab=0x01, line/page/column break=0x02/0x03/0x04, footnote/endnote ref=0x05/0x06, image=0x07 (caller then appends the image hash bytes), opaque=0x0F (caller then appends canonical hash bytes).
- `AppendHash(IrHash)` — raw 32 bytes.
- `AppendStructure(byte marker)` — writes `0x02, marker` (row=0x10, cell=0x11).
- `Build() → IrHash`.

**Format fingerprint** (spec §6.2): `FingerprintRunFormat(IrRunFormat)` / `FingerprintParaFormat(IrParaFormat)` — serialize each non-null field as `name=value;` pairs in declaration order into UTF-8, append `UnmodeledDigest` bytes, SHA-256. `FingerprintBlock(IrParaFormat paraFormat, IEnumerable<IrRunFormat> runFormats)` — paragraph fingerprint bytes + each run fingerprint's bytes in order, hashed.

**Tests:**
- `Canonicalize_AttributeOrder_Irrelevant` — same element, attributes written in different orders → equal bytes.
- `Canonicalize_RsidAndPt14_Stripped` — adding `w:rsidR` / pt14 attrs doesn't change hash.
- `Canonicalize_ContentChange_Detected` — different text → different hash.
- `ContentHash_TextVsSentinel_NoCollision` — `AppendText("a") + AppendSentinel(tab)` ≠ `AppendText("a	…")`-style lookalikes; specifically a literal text "a" followed by tab sentinel differs from text "a" + text containing no tab, and from `AppendText("a")`-free constructions (verify builder output bytes directly).
- `Fingerprint_NullFieldsOmitted` — `IrRunFormat` with only Bold=true vs Bold=true+Italic=null → equal; Bold=true vs Italic=true → different.
- `Fingerprint_UnmodeledDigest_Participates` — same fields, different digests → different fingerprints.

**Steps:** failing tests → red → implement → green → full suite → commit `feat(ir): add canonicalization and hashing (IrHasher)`.

---

## Task 5: `IrReader` — OOXML → IR for the body scope

**Files:**
- Create: `Docxodus/Ir/IrReader.cs`
- Create: `Docxodus/Ir/IrReaderOptions.cs`
- Modify: `Docxodus/WmlToMarkdownConverter.cs:537` — change `private static string? KindFor(XElement el)` to `internal static string? KindFor(XElement el)` (single source of kind resolution; no other change to that file).
- Test: `Docxodus.Tests/Ir/IrReaderTests.cs`

**`IrReaderOptions`** per spec §5.1 (`RevisionView` enum `{ Accept, Reject, FailIfPresent }`, default `Accept`; `IrScopes` flags enum `{ Body, HeadersFooters, Notes, Comments, All }` — only `Body` honored in M1.1, others accepted and ignored).

**`IrReader.Read(WmlDocument doc, IrReaderOptions? options = null)`:**

1. Copy: `new WmlDocument(doc)` so the caller's bytes are never mutated.
2. RevisionView (N13): `FailIfPresent` → scan for `w:ins|w:del|w:moveFrom|w:moveTo|w:rPrChange|w:pPrChange` and throw `DocxodusException` if found; `Accept`/`Reject` → `RevisionProcessor.AcceptRevisions(copy)` / `RejectRevisions(copy)` (see `Docxodus/RevisionProcessor.cs` for exact API; tests at `Docxodus.Tests/WmlComparerTests.cs:758,777` show usage).
3. Open via `OpenXmlMemoryStreamDocument`/`GetWordprocessingDocument`, get body root XDocument, run `UnidHelper.AssignToAllElementsDeterministic(root)` (`Docxodus/UnidHelper.cs:89`).
4. Walk `w:body` children in order:
   - `w:p` → `IrParagraph`. Kind: map `WmlToMarkdownConverter.KindFor(el)` token through `IrAnchor.KindFromToken` (null → treat as `P`). Inlines: walk run-level content —
     - `w:r`: map `w:rPr` → `IrRunFormat` (modeled fields per spec §7.4; every *unmodeled* rPr child element collected into a synthetic container element and hashed via `IrHasher.Canonicalize` → `UnmodeledDigest`; empty leftover → digest of empty bytes, computed once and cached). Inside the run: `w:t` → text (exact, honoring `xml:space`, N11); `w:tab` → `IrTab`; `w:br` → `IrBreak` (`w:type` attr: default Line, `page`→Page, `column`→Column) (N6); anything else inside the run (e.g. `w:drawing`, `w:sym`, `w:fldChar`, `w:instrText`) → `IrOpaqueInline` with canonical hash.
     - Direct paragraph children that aren't `w:r`/`w:pPr` (hyperlinks, fields, sdt, bookmarks, comment plumbing) → one `IrOpaqueInline` each, canonical-hashed. (N3/N7–N9/N12/N14/N15 are M1.2 — opaque is the M1.1 contract.)
     - Drop runs that end up with no inlines and empty text (N10); coalesce adjacent `IrTextRun`s with equal `IrRunFormat` (N5, applied last).
   - `w:p`'s `w:pPr` → `IrParaFormat` (same modeled/unmodeled split; `numPr` present → `IrListInfo(numId, -1, ilvl, "", null, false)` placeholders per scope-reduction note).
   - `w:tbl` → `IrTable`: rows → `IrRow`, cells → `IrCell` (`w:gridSpan` → GridSpan default 1; `w:vMerge` → `IrVMerge` — element absent=None, `@w:val="restart"`=Restart, otherwise Continue); cell children recurse through the same block walker (nested tables work for free). `tblPr`+`tblGrid` canonical-hashed into `UnmodeledTablePropsDigest`.
   - `w:sectPr` (body-trailing) → `IrSectionBreak` (`IrSectionFormat` from `w:pgSz`/`w:pgMar`/`w:type`; leftover children → `UnmodeledDigest`).
   - Any other body child → `IrOpaqueBlock` (canonical hash; anchor kind `Unk`, Unid from the element's `pt:Unid`).
5. Hashes: compute each block's `ContentHash` via `IrContentHashBuilder` (§6.1 — paragraph: inlines in order; table: per-row/per-cell structure sentinels + child block hash rollup) and `FormatFingerprint` via `IrHasher.FingerprintBlock`. Tables' fingerprint = hash of (`UnmodeledTablePropsDigest` + each row's cells' child block fingerprints in order).
6. Anchors: `new IrAnchor(kind, "body", unid)` — scope name `"body"` matching the projection. Build the `AnchorIndex` dictionary over all blocks recursively (tables, rows, cells, and nested blocks all included). Duplicate anchor string → throw `DocxodusException` (invariant).
7. Assemble `IrDocument` with `IrScope("body", blocks)`, empty stores/registries, `Sources` = { body part URI → its XDocument }.

**Tests** (build documents programmatically — see `Docxodus.Tests/DocxSessionTests.cs` and CLAUDE.md note about required parts; a minimal valid docx helper likely already exists in the test project — reuse it; if not, create `Docxodus.Tests/Ir/IrTestDocuments.cs` helper):
- `Read_SimpleParagraphs_ProducesParagraphBlocks` — 2 paragraphs → 2 `IrParagraph` with correct text, anchors with kind `P`, scope `body`, 32-char unids.
- `Read_DoesNotMutateInput` — input byte array identical before/after.
- `Read_Twice_IdenticalAnchorsAndHashes` — determinism: same bytes → same anchor strings, same `ContentHash`/`FormatFingerprint` hex, node-for-node record equality.
- `Read_BoldRun_MapsRunFormat` — bold run → `IrRunFormat.Bold == true`; identical adjacent formatting coalesces into one run (N5).
- `Read_TabAndBreak_BecomeTypedInlines` — `w:tab`/`w:br type="page"` → `IrTab`/`IrBreak(Page)`.
- `Read_Table_StructureAndAnchors` — 2×2 table → `IrTable` with row/cell anchors (`tr`/`tc` kinds) all resolvable via `FindByAnchor`.
- `Read_NestedTable_Recurses` — table-in-cell shape asserted.
- `Read_UnknownElement_BecomesOpaque` — e.g. a `w:sdt` body child → `IrOpaqueBlock`; a `w:hyperlink` in a paragraph → `IrOpaqueInline`.
- `Read_ContentHash_IgnoresFormatting` — same text, one bolded → equal `ContentHash`, different `FormatFingerprint`.
- `Read_RevisionView_AcceptVsReject` — doc with one tracked insertion: Accept-read contains inserted text, Reject-read doesn't; `FailIfPresent` throws.
- `Read_TrailingSectPr_BecomesSectionBreak` — last block is `IrSectionBreak` with page size populated.

**Steps:** failing tests → red → implement (this is the largest task; implement walker incrementally against the tests) → green → full suite → commit `feat(ir): add IrReader for body scope with opaque fallback`.

---

## Task 6: Diagnostic JSON + corpus totality + golden snapshots

**Files:**
- Create: `Docxodus/Ir/IrDiagnosticJson.cs`
- Create: `Docxodus.Tests/Ir/IrCorpusTests.cs`
- Create: `Docxodus.Tests/Ir/Snapshots/` (golden files, committed)
- Test: `Docxodus.Tests/Ir/IrDiagnosticJsonTests.cs`

**`IrDiagnosticJson.Write(IrDocument) → string`** (spec §9): stable, indented JSON via `System.Text.Json.Utf8JsonWriter` — NOT reflection serialization (stability is the requirement). Per block: `anchor`, `type` (`paragraph|table|sectionBreak|opaque`), `contentHash` (hex), `formatFingerprint` (hex), then type-specific: paragraphs → `inlines` array (`{kind: "text", text, format: {…non-null fields only…}}`, `{kind: "tab"}`, `{kind: "break", breakKind}`, `{kind: "opaque", element, hash}`), tables → `rows`/`cells` recursion, sectionBreak → format fields, opaque → `element` name + hash. Dictionary-order independence: emit format fields in declaration order, hex lowercase, no timestamps, no machine paths.

**Corpus totality test** `Read_EntireTestFilesCorpus_DoesNotThrow`:
- Enumerate `TestFiles/**/*.docx` (resolve path relative to test assembly as existing tests do — check how existing tests locate `TestFiles/`, e.g. in `Docxodus.Tests/TestUtil.cs` or similar, and reuse that mechanism).
- For each file: try `new WmlDocument(path)` + a plain `WordprocessingDocument` open; if *that* fails, skip (file is an intentionally-broken fixture — record as skipped).
- Else `IrReader.Read` must succeed; collect all failures and `Assert.Empty(failures)` with file names + exception messages in the failure output.

**Golden snapshots** `Read_CuratedFixtures_MatchGoldenSnapshots`:
- Curated list of ~8 fixtures spanning features, chosen from `TestFiles/` at implementation time (at minimum: one simple multi-paragraph doc, one with tables, one with tracked changes (read under Accept), one with lists/numbering, one with images, one with footnotes, one header/footer doc, one large/complex doc — pick by inspecting the directory; document the choice in a comment).
- For each: `IrDiagnosticJson.Write(IrReader.Read(...))` compared byte-for-byte to `Docxodus.Tests/Ir/Snapshots/<fixture>.ir.json`.
- Regeneration affordance: environment variable `DOCXODUS_IR_REGEN_SNAPSHOTS=1` rewrites the files instead of asserting (document in a comment at the top of the test class).
- `Snapshots/*.ir.json` files marked `CopyToOutputDirectory` in `Docxodus.Tests.csproj` following however `TestFiles/` is wired (check and mirror).

**Determinism test** `DiagnosticJson_TwoReads_ByteIdentical` — read the same fixture twice, JSON strings equal.

**Steps:** failing tests → red → implement writer → generate + commit snapshots (review them by eye first: anchors look right, hashes present, no absolute paths) → green → full suite → commit `feat(ir): add diagnostic JSON, corpus totality test, golden snapshots`.

---

## Self-review checklist (controller, after Task 6)

- All spec §-references implemented or explicitly deferred with an M1.2/M1.3 pointer in code comments.
- `dotnet build -c Release Docxodus.sln` clean (warnings-as-errors).
- `./scripts/build-wasm.sh` still succeeds (IR must not reference SkiaSharp; run `dotnet clean` after, per CLAUDE.md).
- CHANGELOG.md `[Unreleased]` entry added (internal feature — brief note under `### Added`).
- No public API surface added (everything `internal`).
