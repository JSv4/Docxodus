# Document IR

> **Status:** Phase 1 (M1.1–M1.4) complete on `feat/document-ir`. **Internal and
> experimental** — every type under `Docxodus/Ir/` is `internal`; there is no
> public API, WASM/npm/python surface, or stability guarantee yet. The IR exists
> to be consumed by future engines (the Phase 2 diff engine, a deferred layout
> engine) and is validated today by one consumer: the IR markdown emitter
> (`IrMarkdownEmitter`), which reproduces the shipped markdown projection's
> output. Sources of truth that this doc points at rather than duplicates: the
> [IR spec](../superpowers/specs/2026-06-11-document-ir-spec.md) (detailed type
> model + the N1–N15 normalization table + hashing) and the
> [program plan](../superpowers/specs/2026-06-11-ir-diff-layout-program-plan.md).

## Overview & Motivation

Today every major Docxodus module re-derives its own private, throwaway view of
a document from raw OOXML: `WmlComparer` builds `ComparisonUnitAtom` lists,
`WmlToMarkdownConverter` builds its projection, `OpenContractExporter` builds a
text/structure view, `FormattingAssembler` resolves the style cascade
destructively on the XML. Each re-implements normalization with subtly different
rules — which is the root of "module A and module B disagree about whether these
two paragraphs are equal" bugs.

The Document IR replaces those private views with **one shared, immutable,
typed, normalized, anchor-identified in-memory model of a DOCX**, built once per
open and consumed many times:

```
                       ┌────────────────────────────┐
      OOXML (.docx) ──►│         Document IR        │──► markdown projection (this phase's validator)
                       │  typed · normalized ·      │──► diff engine        (Phase 2)
                       │  anchor-identified ·       │──► layout engine      (Phase 3, deferred)
                       │  immutable snapshot        │
                       └────────────────────────────┘
```

The IR is **read-only** (no IR→OOXML writer in Phase 1), **lossy-tolerant**
(anything unmodeled becomes an `Opaque` node that still hashes correctly), and
**deterministic** (two reads of the same bytes produce node-equal, hash-equal,
anchor-equal snapshots — a tested invariant, not an aspiration).

Entry point: `IrReader.Read(WmlDocument, IrReaderOptions)` → `IrDocument`.

## Type Model (summary)

Full shapes live in spec §7; the namespace is `Docxodus/Ir/`. The essentials:

| Type | Role |
|---|---|
| `IrDocument` | Root snapshot: `Body` scope, per-part `Headers`/`Footers`, `Footnotes`/`Endnotes` note stores, `Comments` store, the three registries, and `Sources` (part URI → `XDocument`, the provenance pin). `FindByAnchor` is O(1). |
| `IrScope(Name, Blocks)` | A flat block list for one part. `Name` matches the projection's multipart namespacing: `body`, `hdr1…N`, `ftr1…N`, `fn`, `en`, `cmt`. |
| `IrBlock` (abstract) | Base for `IrParagraph`, `IrTable`, `IrSectionBreak`, `IrOpaqueBlock`. Carries `Anchor`, `ContentHash`, `FormatFingerprint`, and an equality-excluded `Source`. |
| `IrParagraph` | `Format` (direct) + lazy `EffectiveFormat`, optional `List` facts, `Inlines`, and the rendering-derived `ResolvedListMarker` (see below). |
| `IrTable` / `IrRow` / `IrCell` | Table tree; cells carry `GridSpan`/`VMerge`; rows/cells/tables each roll up a `ContentHash`. |
| Inlines | `IrTextRun`, `IrTab`, `IrBreak`, `IrHyperlink`, `IrFieldRun(Instruction, CachedResult)`, `IrNoteRef`, `IrInlineImage`, `IrOpaqueInline`. Runs are **not** anchored — inline positions are addressed as (block anchor, char span). |
| Formats | `IrParaFormat`, `IrRunFormat`, `IrSectionFormat` — direct format records, each carrying an `UnmodeledDigest` (§6.4) so unmodeled properties still move the fingerprint. |

All nodes are C# `record`s, so value-equality is structural. Two members are
deliberately kept out of equality: `Source` (provenance — see below) and any
derived-index dictionaries on `IrDocument`/registries.

### `ResolvedListMarker` and the equality footgun

`IrParagraph.ResolvedListMarker` is the auto-number string Word would render
(`"1."`, `"a."`, `"1.1"`, a bullet glyph, "First Article"). It is resolved by the
reader against the live package (Word numbering is a stateful document-order
counter walk that the per-paragraph `List` facts cannot capture), so it is an IR
fact, not an emitter computation. Because it is an init-only record member it
**participates in record value-equality, making record `==` stricter than
`ContentHash` equality** — the marker is rendering-derived state, not content.
**Phase-2 diff code must key alignment/equality on the hashes, never on record
`==`.** (Noted in the XML-doc on the member itself.)

## Identity / Anchors

Identity is `IrAnchor(Kind, Scope, Unid)`, whose `ToString()` is the markdown
projection's grammar: `kind:scope:unid` → `{#p:body:a1b2c3d4}`.

- **Unids** come from the existing deterministic content-addressable assignment
  pipeline (`UnidHelper.AssignToAllElementsDeterministic`, the same code the
  `DocxSession` open path uses) — the reader **reuses** it, it does not reinvent
  it. A Unid is a SHA-256 of `parent_unid : tag : content_signature :
  duplicate_index`, so the same bytes yield the same anchors across sessions and
  across code paths.
- **Kind resolution** (`p` vs `h` vs `li`) follows the projection's outline-level
  / list-membership rules so the same element gets the same anchor string from
  both the IR and the shipped projection. **M1.4 asserts string equality of
  anchors between the two paths** — non-negotiable, since `DocxSession` clients
  hold these ids.
- **Anchored:** paragraphs, tables, rows, cells, section breaks, footnotes,
  endnotes, comments, images/drawings, opaque blocks. **Not anchored:** runs and
  other inlines (run identity is unstable across edits by nature).

The anchor scheme is documented in full in
[`markdown_projection.md`](./markdown_projection.md#anchor-scheme); the IR adopts
it verbatim so the two agree.

### Provenance — `IrProvenance`

Every node carries `Source` (`IrProvenance`): the originating `XElement?` and the
owning `PartUri?`, plus an equality-neutral positional fact `FromBlockSdt` (true
when a block-level content-control wrapper was unwrapped to deliver this block).
`IrProvenance.Equals` returns `true` for any other instance and `GetHashCode`
returns `0`, so provenance is structurally excluded from record equality: two IR
trees from different physical files with identical content compare equal. The
`IrDocument` pins the parsed `XDocument`s in `Sources` so provenance pointers
stay alive exactly as long as the snapshot (the memory consequence is in the
budget below).

**Optional retention — `IrReaderOptions.RetainSources` (default `true`).** Pinning
the XML DOM is convenient for diagnostics/raw round-tripping but roots ~11× the part
XML for the snapshot's lifetime. The Phase-2 diff engine and bulk pipelines don't
need element-level provenance, so retention is opt-out: with `RetainSources=false`,
`Sources` is empty and every node's `Source` is a single shared empty `IrProvenance`
(`Element`/`PartUri` null, zero per-node allocation), letting the working
`XDocument`s be collected once `Read` returns (≈11× → ≈2.7× retained — see the gate
addendum). The part URI a consumer needs at scope/block granularity survives as a
**scope-level** fact, promoted to `IrScope.PartUri` (and `IrCommentStore.PartUri`),
populated in BOTH modes; the markdown emitter prefers those over per-node provenance.
Retention is a pure memory knob — anchors, `ContentHash`, and `FormatFingerprint` are
byte-identical across modes.

## Normalization

The reader applies normalization rules **N1–N15** before any node is
constructed; the spec's table (§5.2) is the **source of truth** and each rule has
a dedicated unit test (`IrNormalizationTests`). Highlights: strip `rsid*`
(N1) and proofing/layout-cache noise (N2/N4); coalesce equal-format adjacent
runs (N5); `w:tab`/`w:br` become structural inlines, never folded into text
(N6); `w:noBreakHyphen`→U+2011 / `w:softHyphen`→U+00AD (N7); fields (`w:fldSimple`
and complex `fldChar` runs) become `IrFieldRun(Instruction, CachedResult)` (N9);
text preserved exactly with `xml:space`, no whitespace conflation — conflation is
a *diff-time* setting, not an IR fact (N11); SDT/smartTag unwrapped to content
with the anchor on the outer block (N12); revisions resolved per `RevisionView`
before node construction (N13); comment plumbing recorded as (block anchor, char
span) in the comment store, never affecting `ContentHash` (N15).

## Hashing

Every block carries two hashes (`IrHash`, 32-byte SHA-256). The pair is the diff
engine's primary signal: **equal/equal → unchanged; equal/different → format-only
change; different → content change.**

- **`ContentHash`** — text identity (what a reader sees), computed over a
  canonical UTF-8 byte stream (spec §6.1). Non-text inlines contribute a
  **sentinel byte sequence** `0x01 <kind-byte>` (tab=0x01, breaks=0x02/03/04,
  note refs=0x05/06, image=0x07 + image part hash, textbox=0x0B + each inner
  block's content hash, opaque=0x0F + canonical hash)
  — sentinels live outside the Unicode text range so text can never collide with
  structure. Hyperlinks frame their target between `0x08`/`0x09` sentinels so
  linked text is never content-equal to identical plain text and a target change
  reads as a content change. `IrFieldRun` hashes its *cached result* unbracketed
  (a field rendering "5" is content-equal to a literal "5" — deliberate); the
  instruction is consumer-visible but **unhashed**. `IrNoteRef` hashes its kind
  sentinel **without** the note id, so renumbering alone never flips body hashes.
  Formatting never affects `ContentHash`.
- **`FormatFingerprint`** — hash of the node's **direct** (not effective) format
  record plus its `UnmodeledDigest`; run fingerprints roll up into the block
  fingerprint with the paragraph's own `IrParaFormat`. Deliberately direct: a
  style-definition edit should read as one style change, not N paragraph edits
  (the diff engine compares style definitions separately).
- **`UnmodeledDigest`** (§6.4) — when the reader maps `w:rPr`/`w:pPr` into format
  records, any child it does not model is canonicalized (§6.3: attributes sorted,
  bookkeeping/rsid stripped, whitespace normalized) into this digest, so a change
  in an unmodeled property still flips the fingerprint instead of silently
  reading as equal. Lossy-tolerance applied to formatting.
- **Opaque canonicalization** (§6.3) gives `IrOpaqueBlock`/`IrOpaqueInline` a
  `CanonicalHash` stable across attribute reordering and rsid churn, so unmodeled
  content still diffs as unchanged/changed/moved without being understood.

## Registries

Resolved once during `Read`, before the body walk, so paragraph list resolution
can chase `numPr → IrNum → IrAbstractNum → level format` (`Docxodus/Ir/IrRegistries.cs`):

- **`IrStyleRegistry`** — styles by id (with `BasedOn`, cloned `w:pPr`/`w:rPr`),
  the default paragraph style id, and doc-defaults.
- **`IrNumberingRegistry`** — `IrNum` (numId → abstractNumId + start-overrides)
  and `IrAbstractNum` (abstractNumId → per-ilvl `IrNumLevel`).
- **`IrThemeFonts`** — major/minor ASCII theme fonts.

Tolerant population: a missing part resolves to the registry's `Empty` value;
malformed entries are skipped, not fatal; duplicate ids are first-wins.
Registries hold cloned `XElement`s and dictionaries that compare by reference, so
they are **derived indexes excluded from content-scope determinism** — callers
must not rely on registry value-equality for document equality. (`numStyleLink`
indirection is deferred past M1.3.)

## Effective Formats

`IrParagraph.EffectiveFormat` (and the run-level equivalents) are **lazy,
cascade-resolved, cached** projections that walk doc-defaults → style chain
(`BasedOn`) → direct properties, reusing `FormattingAssembler`'s *logic*
**non-destructively** (the IR does not call the destructive assembler).
Resolution is semantically pure, so the laziness is unobservable and the snapshot
stays freely shareable across threads. Effective formats exist for consumers (the
projection, future layout); the diff engine deliberately uses **direct** formats
for fingerprints (see Hashing). Parity with `FormattingAssembler`/
`GetListMembership`/`GetBlockMetadata` is spot-checked in tests (M1.3).

## Scopes

`IrDocument` exposes every part as a scope whose `Name` matches the projection's
namespacing so anchors agree: `Body`, per-part `Headers`/`Footers` (`hdr1…N`/
`ftr1…N`, each tagged default/first/even with section linkage), the
`Footnotes`/`Endnotes` note stores (note id → `IrScope`, boilerplate
separator/continuation notes filtered by consumers), and the `Comments` store
(comment ranges as (block anchor, char span) per touched block).

## The Markdown Emitter & Equivalence Status

`IrMarkdownEmitter.Emit(IrDocument, WmlToMarkdownConverterSettings)` is the
Phase 1 validator: an **IR consumer** that reproduces the shipped
`WmlToMarkdownConverter`'s output (markdown string + `AnchorTarget` index). The
shipped converter stays the byte-untouched **oracle**; the emitter never mutates
it. It ports headings/plain/list lines with `{#kind:scope:unid}` anchors, the
inline formatting subset (bold/italic/code/strike/links with the oracle's exact
delimiters and escaping), tabs/breaks, GFM-vs-opaque tables, image lines, opaque
fenced summaries, section thematic breaks, multipart `# Headers`/`# Footers`/
`# Footnotes`/`# Endnotes`/comments sections, auto-number prefixes (from
`ResolvedListMarker`), and the `EmptyParagraphs`/`AnchorIdRendering` settings
modes.

**Equivalence: 642/668 corpus fixtures byte-equal** (markdown + body anchor
index; 608/668 at the Phase-1 gate, lifted to 642 by M1.5 textbox modeling). The
remaining divergences are controller-adjudicated; the dispositions and the full
triage table are in the
[M1.4 gate report](../superpowers/plans/2026-06-11-ir-m14-markdown-projection-port.md#outcome-phase-1-gate-report).
Two classes are **accepted** divergences (oracle bugs — the IR output is *more*
correct): special-character drops and multi-run hyperlink splits. The rest are
**deferred** small-fixture-sweep IR work (CC/revision spacing, heading-numPr
display, TOC field text). The former dominant gap — the **textbox/shape-body
gap** — is now closed: textbox bodies are modeled as `IrTextbox` (inner blocks
anchored/hashed/indexed; sentinel `0x0B`), so both the markdown body and the
header/footer `ScopeHasContent` content-detection gate match the oracle's raw
`w:t` view. The harness (`IrMarkdownEquivalenceTests`,
Trait `Corpus`) drives both paths over the corpus, writes per-fixture diffs to
the gitignored `EquivalenceArtifacts/`, and is the loop driver, not just a
pass/fail gate. The perf budget lives in `IrMarkdownPerfBudgetTests` (Trait
`Perf`); its full corpus benchmark is opt-in via `DOCXODUS_RUN_PERF=1` (it forces
blocking GCs that would flake concurrent SkiaSharp image tests in the default
parallel run), with a GC-quiet smoke check as the default-run guard.

## Evolution Policy

- **Opaque promotion** (modeling a previously-opaque element) is additive but
  changes hashes and golden snapshots for affected documents: it lands with
  regenerated-and-**reviewed** snapshots (never blind-regenerated — every diff is
  triaged) and a CHANGELOG entry.
- **Hashes and Unids are session-scoped identities** (like Unids): consumers must
  **never persist IR hashes across library versions**.
- **Visibility:** types go `public` only when Phase 2 productization needs them,
  under a documented "experimental" banner.
- **v2 candidates, explicitly deferred:** revision-aware IR (`w:ins`/`w:del` as
  nodes for as-is tracked-changes projection), bookmarks as ranges,
  content-control metadata as typed facts, and an IR→OOXML writer beyond the
  Phase 2 renderer's needs. (Textbox/shape body content is now modeled — see
  `IrTextbox`, M1.5 — and is no longer deferred.)

## Current Limitations

- **Textbox bodies are modeled inline, not as a separate scope.** `IrTextbox`
  carries the inner blocks at the containing paragraph's inline position (anchored
  in the containing scope, hashed, indexed) rather than as a distinct
  textbox/shape scope. This matches the oracle's `Descendants`/`DescendantsAndSelf`
  view and closes the `ContentHash` blind spot; a dedicated shape scope (with
  shape geometry/anchoring facts) remains possible future work.
- **No as-is projection of tracked-changes documents** — `RevisionView` defaults
  to Accept; an as-is (revision-aware) view is a v2 item. The equivalence harness
  pre-accepts revisions once and feeds the same bytes to both paths to compare
  like-for-like.
- **Memory footprint is above the ≤3× reference target in the default (retained)
  mode.** A retained snapshot costs roughly (pinned XML DOM via `Sources`) + (IR
  nodes); the largest-body corpus fixture retains ≈11× its main-part XML size
  (measured, reported — see the gate report). M1.5 made this opt-out:
  `RetainSources=false` drops the pinned DOM and brings the same fixture to ≈2.7×
  (see Provenance above + the gate addendum), which is the mode Phase-2 bulk
  consumers should use.
- **`numStyleLink` numbering indirection** and several v2 model facts above are
  not yet resolved.

## Relationship to the Spec & Plans

- **Detailed design / source of truth:**
  [`specs/2026-06-11-document-ir-spec.md`](../superpowers/specs/2026-06-11-document-ir-spec.md)
  (type model §7, normalization table §5.2, hashing §6, conformance/budgets §10,
  evolution §11).
- **Program plan & decision log:**
  [`specs/2026-06-11-ir-diff-layout-program-plan.md`](../superpowers/specs/2026-06-11-ir-diff-layout-program-plan.md).
- **M1.4 task plan + Phase 1 gate report:**
  [`plans/2026-06-11-ir-m14-markdown-projection-port.md`](../superpowers/plans/2026-06-11-ir-m14-markdown-projection-port.md).
- **The validated consumer's own spec:**
  [`markdown_projection.md`](./markdown_projection.md).
