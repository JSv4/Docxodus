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

## M2.1 Outcome

**Status: COMPLETE (2026-06-11).** All three tasks landed; exit criteria met. The diff layer (`Docxodus/Ir/Diff/`, all `internal`, `#nullable enable`, WASM-safe, no new dependencies) holds `IrDiffSettings`/`IrDiffToken`/`IrDiffTokenizer` (Task 1), `IrBlockAlignment`/`IrBlockAligner` (Task 2), and the corpus/adversarial/scale coverage below (Task 3). The shared `IrAlignmentAsserts` helper (extracted from Task 2's inline invariants) is reused by every aligner test.

### WC corpus alignment smoke (`IrAlignerCorpusTests`, `Trait Category=Corpus`)

- **92 base↔variant pairs** inferred from `TestFiles/WC/` by name convention (rules documented in the test's class doc + `BuildPairs`): (1) `-Before…` ↔ `-After…` families split at the first Before/After token, with index-matched pairing for multi-base families (`WC021 Before-1/After-1`, `Before-2/After-2`) and single-before fan-out to every numbered after (`WC033`/`WC034` `Before` ↔ `After1/2/3`); (2) base ↔ prefix-extending variants (`WC001-Digits` ↔ `…-Mod` AND ↔ `…-Deleted-Paragraph`; `WC006-Table` ↔ both row-delete variants); (3) `WCnnn-` numeric-prefix fan-out around an `-Unmodified` base (`WC002`, `WC007`). **161 of 163 WC files** are covered; the two `WC014-SmartArt-With-Image-Deleted-After[2]` files are deliberately unpaired (the `-Deleted-` infix gives no unambiguous base; the family's plain Before/After/After2 pairs exercise the same content).
- **Every pair runs forward (before→after) AND reversed (after→before)**; the shared invariants (totality by reference identity, per-kind hash constraints, `MovedModified` never produced) hold in both directions — **no throws, all invariants pass**.
- **Per-pair kind histograms + corpus totals** logged via `ITestOutputHelper`. Corpus totals (forward): Unchanged=556, FormatOnly=1714, Modified=1488, Moved=3, MovedModified=0, Inserted=901, Deleted=35. Highlights: most small WC pairs resolve to a couple of Unchanged + one Modified (the edited block), exactly as expected; tables/SmartArt/images align as whole-block Modified units (M2.1 granularity); the large `WC-BodyBookmarks-Before/After` pair (the only Moved source) yields FormatOnly=1714/Modified=1374/Moved=3 with Inserted/Deleted swapping cleanly under reversal (885↔28).

### Adversarial fixtures (`IrAlignerAdversarialTests`)

- **500 near-identical** distinct-clause paragraphs, one word changed in one → **499 Unchanged + 1 Modified, 0 Moved** (0 Inserted/Deleted). ✓
- **500 identical** boilerplate paragraphs, one deleted → **499 Unchanged + 1 Deleted, 0 Moved, 0 Modified** (the gap in-order refinement resolves repeated content without false moves). ✓
- **Fully rewritten** 200 vs 200 completely-different paragraphs → **200 Modified, 0 Unchanged, 0 Moved**, no throw, sub-second runtime (one head↔tail gap, all positional Modified). ✓
- **Contiguous block move**: a 10-paragraph block of unique paras relocated front→back of a 300-para doc → **exactly 10 Moved + 290 Unchanged**. This confirms the LIS spine keeps the 290 stationary blocks (right positions 0..289, left 10..299, monotone) and drops the smaller 10-block off the spine as the move — the moved 10 (left 0..9) cannot extend the spine past the stationary chain. The smaller side dropping off is the correct, designed behavior.

### Scale guard (`Trait Category=Perf`, default-run-safe)

- 500-para vs 2000-para near-identical self-pairs (one edit each), warm-up + best-of-3: **500 = 1.40 ms, 2000 = 6.62 ms → 4.72× for 4× input** (≤ 8× anti-O(n²) bound). Inputs are sized to all-unique anchors so no single large all-distinct gap trips the `InOrderRefine` G²/2 worst case; this isolates the anchoring/spine cost, which scales near-linearly.

### Known limitations carried to M2.2

- **Cross-gap move + edit → Delete + Insert.** M2.1 move detection is exact-`ContentHash` only. A block that is BOTH moved and edited has no exact off-spine anchor, so it falls out as Deleted + Inserted (or a positional Modified if it happens to land in a shared gap), never as a move. Similarity-based fuzzy move matching is M2.2.
- **`MovedModified` reserved, never produced.** The enum kind exists for surface stability but M2.1 cannot reach it (it needs intra-block token diff + fuzzy move pairing).
- **Whole-block table granularity.** Tables align as single units; a cell-only edit surfaces as one Modified table entry. Row/cell-level alignment is M2.2+.
- **Positional gap pairing.** Non-anchored blocks inside a gap pair by order, not similarity; similarity-based gap-pairing refinement is M2.2.
- **FormatFingerprint over-sensitivity to run-boundary churn (review finding on the WC corpus run).** 100% of the corpus' 1,714 FormatOnly entries come from one heavily-edited 2 MB real-world pair (`WC-BodyBookmarks`), where content-identical paragraphs register as FormatOnly because editing churn re-segments runs with per-run rPr noise (surviving in `UnmodeledDigest`, defeating N5 coalescing), flipping the ordered run-format sequence the block fingerprint digests. Not rsid (stripped). M2.2 should consider a boundary-normalized or order-insensitive run-format digest — and/or additional rPr-noise normalization rules — before FormatOnly classifications drive `w:rPrChange` emission.
