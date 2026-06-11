# Diff Engine — M2.1 Tokenizer + Block Alignment

> **For agentic workers:** REQUIRED SUB-SKILL: superpowers:subagent-driven-development. Executed by the controller as sequenced subagent tasks.

**Goal:** Phase 2 opens. A diff-time tokenizer over IR paragraphs and a block-level alignment engine with move detection integrated into the alignment (unique-hash anchoring + LIS spine), producing the typed alignment that M2.2's intra-block diff and edit script consume.

**Baseline:** `feat/diff-engine` @ 9f81c57 (main post-M1.5: IR hardened, 1.16× read cost, `RetainSources=false` available, no ContentHash blind spots).

**Program-plan contract (M2.1):** tokenizer honoring `WordSeparators`/culture/case as *diff settings*; alignment over `ContentHash`/`FormatFingerprint` pairs using unique-hash anchoring (histogram-diff style) with moves falling out of alignment by construction; adversarial fixtures (500 near-identical paragraphs, boilerplate-heavy) with complexity assertions.

**Layout:** new `Docxodus/Ir/Diff/` folder, namespace `Docxodus.Ir.Diff`, all `internal`, `#nullable enable`, WASM-safe, no new dependencies. Reads IR built with `RetainSources = false`.

## Task 1: Diff settings + tokenizer

`IrDiffSettings` (record/class): `WordSeparators` (default = `WmlComparerSettings`' default set — copy the values, cite the source), `CaseInsensitive` (false), `ConflateBreakingAndNonbreakingSpaces` (true — matches WmlComparer's default), `CultureInfo?` (null → ordinal). `IrDiffTokenizer.Tokenize(IrParagraph, IrDiffSettings) → IReadOnlyList<IrDiffToken>`:

```csharp
internal enum IrDiffTokenKind { Word, Separator, Tab, Break, NoteRef, Image, FieldResultBoundary—NO (fields are transparent), Opaque, Textbox, HyperlinkBoundary—NO (decide below) }
internal sealed record IrDiffToken(IrDiffTokenKind Kind, string Text, string MatchKey,
                                   int StartChar, int EndChar, IrRunFormat? Format, IrHash? AtomHash);
```

Semantics (mirror the §6.1 content-hash stream so token equality ⇔ content-hash equality at the same granularity):
- Text runs split on `WordSeparators` into Word + Separator tokens; `MatchKey` = text after settings normalization (case fold per `CaseInsensitive` + culture; NBSP→space when conflating). `Text` stays raw. Char offsets = the same coordinate space as comment targets / `ApplyFormat` (emitted IrTextRun chars).
- `IrTab`/`IrBreak(kind)`/`IrNoteRef(kind)` → atomic tokens, MatchKey = a sentinel-style key (kind-distinct, id-less for note refs — consistent with hashing).
- `IrInlineImage` → atomic token, MatchKey from `ImageBytesHash`. `IrOpaqueInline` → atomic, MatchKey from `CanonicalHash`. `IrTextbox` → ONE atomic token whose MatchKey is the rolled-up inner-block hash sequence (inner blocks are aligned as blocks separately — the paragraph-level token is just a placeholder; document).
- `IrFieldRun` → its CachedResult inlines tokenized transparently (consistent with N9/§6.1). `IrHyperlink` → child inlines tokenized transparently BUT each token carries the hyperlink target in its MatchKey suffix (consistent with §6.1's framed-target hashing: linked text ≠ plain text; target change = content change). Document both decisions.
- Each token carries the governing `IrRunFormat` (for M2.2 format-change detection).

Tests: word/separator splitting incl. multi-separator runs, case-fold + NBSP settings behavior, offsets line up with the text, atomic kinds, hyperlink-target-in-key, field transparency (PAGE-result "5" tokens == literal "5" tokens), determinism.

## Task 2: Block alignment with integrated moves

Types (`IrBlockAlignment.cs`):

```csharp
internal enum IrAlignmentKind { Unchanged, FormatOnly, Modified, Moved, MovedModified, Inserted, Deleted }
internal sealed record IrAlignedBlock(IrAlignmentKind Kind, IrBlock? Left, IrBlock? Right);
internal sealed record IrBlockAlignment(IReadOnlyList<IrAlignedBlock> Entries);   // document-order (right-side order, with deletions interleaved at their left positions — define + document precisely)
```

`IrBlockAligner.Align(IrDocument left, IrDocument right, IrDiffSettings) → IrBlockAlignment` over the body block sequences (tables align as whole blocks in M2.1; row/cell-level alignment is M2.2+ — document):

1. **Anchor pass (histogram-style):** key = `(ContentHash, FormatFingerprint)`; blocks whose key occurs exactly once on each side pair up. Second pass on `ContentHash` alone for FormatOnly candidates (unique each side).
2. **Spine:** longest increasing subsequence over the anchored pairs' (leftIndex, rightIndex) → in-order spine = Unchanged/FormatOnly. Anchored pairs OFF the spine = **Moved** (or **MovedModified** never arises from exact matches — reserve the kind for M2.2; document).
3. **Gap fill:** between consecutive spine pairs, remaining left/right blocks pair positionally in order → **Modified** candidates (M2.2 runs the token diff inside these); unpaired left → Deleted, right → Inserted. Repeated-content blocks (non-unique hashes — boilerplate) resolve within gaps by order, never globally (this is what keeps boilerplate-heavy docs O(n)).
4. Determinism: same inputs → identical alignment; no randomness, no dictionary-order dependence (sort/stable-iterate everything).

Tests on synthetic IR (build via IrTestDocuments + IrReader): identity alignment (all Unchanged), single edit (Modified), insert/delete, pure move (paragraph relocated → Moved, NOT delete+insert — the headline capability), move+unrelated-edit coexistence, format-only (bolded paragraph → FormatOnly), duplicate-content blocks don't false-move (boilerplate: 10 identical paragraphs, one deleted → one Deleted + 9 Unchanged, zero Moved), table block aligns as unit, empty documents.

## Task 3: Real-pair + adversarial coverage, scale guard, close

- **WC corpus alignment smoke** (`Trait Corpus`): for every TestFiles/WC/ base+variant pair (enumerate by name convention — `*-Before/-After` plus the `WCnnn-X.docx`/`WCnnn-X-Mod.docx` families; build the pair list by inspection, document it): IR-read both (RetainSources=false) + Align; assert totality (no throw), and invariants: every left block appears exactly once across entries; every right block exactly once; Unchanged entries have equal ContentHash+FormatFingerprint; FormatOnly have equal ContentHash only. Output per-pair entry-kind histograms via ITestOutputHelper.
- **Adversarial fixtures** (programmatic): 500 near-identical paragraphs with one word changed in one (expect 499 Unchanged + 1 Modified, no Moved); 500 identical paragraphs with one deleted (boilerplate stress: 1 Deleted, 0 Moved); fully-rewritten document (all Modified/Inserted/Deleted, no pathological runtime).
- **Scale guard** (`Trait Perf`, generous bound): align 500-para and 2000-para versions; assert wall-time ratio ≤ ~8× (order-of-magnitude guard against accidental O(n²); informational numbers logged).
- Close: `## M2.1 Outcome` appended here; program-plan M2.1 marked; CHANGELOG line.

## Out of scope (M2.2+)

Intra-block token diff, edit script, row/cell-level table alignment, similarity-based gap pairing refinement, MovedModified detection, any renderer.
