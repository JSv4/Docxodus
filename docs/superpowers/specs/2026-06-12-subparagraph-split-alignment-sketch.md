# Sub-Paragraph Split/Merge Alignment — Design Sketch (Phase-2 follow-on)

**Date:** 2026-06-12
**Status:** SKETCH ONLY — not scheduled. Produced under the M2.5 Task 2 timebox
rule (a real sub-paragraph alignment model is sketch-and-defer territory, not a
rush job). Recommended as a Phase-2 follow-on item.
**Companion:** [`2026-06-11-ir-diff-layout-program-plan.md`](./2026-06-11-ir-diff-layout-program-plan.md)
(decision log entry M2.5-followon), `docs/architecture/comparison_engine.md`.
**Deviations it closes:** WC-1450, WC-1830 (the two retained GetRevisions
deviations whose root cause is 1:N paragraph split).

## Problem (the established evidence)

`WmlComparer` runs a single whole-document atom-level LCS. The IR diff engine
aligns at **block (paragraph) grain** first (`IrBlockAligner`), then token-diffs
*inside* each 1:1-paired block (`IrTokenDiffer`). When one before-paragraph's
content **migrates across two after-paragraphs** (a paragraph SPLIT — the user
pressed Enter mid-paragraph, or inserted a new paragraph that swallowed the
break), the oracle's flat atom LCS handles it naturally; the IR's block model
**cannot pair one left block with two right blocks**, so it surfaces an extra
whole-paragraph revision.

### WC-1830 — `WC/WC041-Table-5.docx` → `…-Mod.docx`

A single table cell. (Anchors abbreviated.)

| | LEFT cell blocks | RIGHT cell blocks |
|---|---|---|
| 0 | `Video provides … your point. When you click … to add.` (ONE paragraph) | `Video provides … your point. ` (split prefix) |
| 1 | `You can also type a keyword …` | `` (empty math paragraph holding `A=πr2`) |
| 2 | `` (empty) | `When you click … to add.` (split suffix) |
| 3 | — | `` (empty) |

**Oracle GetRevisions (2):** `Deleted "When you click … to add."` +
`Inserted "\nA=πr2\nWhen you click … to add."`. The oracle credits BOTH
`Video provides … your point.` (RIGHT[0]) AND `When you click … to add.`
(RIGHT[2]) as **Equal** against the single LEFT[0] paragraph; only the inserted
paragraph mark + math is a change.

**IR GetRevisions (3, WmlComparer-compatible mode):** the gap-fill similarity
pass pairs LEFT[0] (combined) ↔ RIGHT[2] (`When you click`, the higher-Jaccard
half) as Modified, then RIGHT[0]/RIGHT[1] fall out as Inserted and LEFT[1]
(`You can also type…`) as Deleted. The shared `Video provides…` prefix
(LEFT[0]↔RIGHT[0]) is **lost** because LEFT[0] is already consumed. Net **+1**.

### WC-1450 — `WC/WC023-Table-4-Row-Image-Before.docx` → `…-After-Delete-1-Row.docx`

**The catalog's old description ("two IDENTICAL `Video provides…` cell
paragraphs; the aligner anchored the wrong one") was STALE/WRONG.** The actual
+1 is the SAME 1:N split, inside the surviving table cell `tc:…cd11`:

| | LEFT cell blocks | RIGHT cell blocks |
|---|---|---|
| 0 | `Video provides … your point. When you click … to add.` (ONE paragraph) | `Video provides … your point.` (split prefix) |
| 1 | `You can also type a keyword …` | `` (empty) |
| 2 | `` (empty) | `When you click … to add.` (split suffix) |
| 3 | — | `` (empty) |

Identical mechanism, identical residual (+1). Oracle 7 / IR 8.

## Why no bounded fix exists within the 1:1 model

The crux is provable, not a tuning question:

1. **`IrEditOp` is strictly 1:1** — one `LeftAnchor`, one `RightAnchor`,
   one `IrTokenDiff` over one left paragraph vs one right paragraph
   (`Docxodus/Ir/Diff/IrEditScript.cs`). There is **no shape** that maps one
   left block to two right blocks.
2. The oracle's result requires crediting `When you click` (RIGHT[2]) as Equal
   against LEFT[0]'s SUFFIX while *also* crediting `Video provides`
   (RIGHT[0]) as Equal against LEFT[0]'s PREFIX. That is one paragraph token-
   diffed against TWO — a 1:2 token diff by definition.
3. **Render-time coalescing cannot recover it.** `IrRevisionRenderer`'s
   WmlComparer-compatible coalescing merges *contiguous same-direction whole-
   block ins/del runs* and trims common affixes *within one ModifyBlock's token
   diff*. Here the ops are interleaved (Insert, Insert, Modify, Delete) and the
   Modify already binds LEFT[0] to the wrong single right block; no regrouping
   of 1:1 ops reconstructs a 1:2 Equal crediting.
4. **Re-pairing LEFT[0]↔RIGHT[0] instead of RIGHT[2] does not help** — it is
   symmetric: whichever half LEFT[0] binds to, the *other* half becomes a whole-
   paragraph Insert. 1:1 can keep at most one of the two; the oracle keeps both.

Therefore matching the oracle is **engine-level 1:N split semantics**, a new
edit-script capability — explicitly out of scope to ship under the M2.5
timebox ("Do NOT ship new edit-script op kinds under time pressure").

## Proposed design (the follow-on)

### Detection (in gap fill, `IrBlockAligner.FillOneGap`)

After the similarity pass leaves leftovers, before the 1×1-residue rule:
for each unpaired left block `L`, scan **maximal runs of ADJACENT unpaired right
blocks** `R[a..b]` (ignoring empty/whitespace-only paragraphs as connective
tissue, since a split inserts an empty paragraph-mark carrier). Compute
**containment** of `L`'s token multiset by the UNION of the run's multisets:

```
containment(L, R[a..b]) = |bag(L) ∩ (bag(R[a]) ⊎ … ⊎ bag(R[b]))| / |bag(L)|
```

A split candidate fires when:
- `b > a` (a genuine 1:N, N≥2 non-empty right blocks), AND
- `containment ≥ SplitContainmentThreshold` (high, e.g. 0.9 — `L`'s words are
  almost entirely present, in order, across the run), AND
- the union does NOT contain large material foreign to `L` beyond a bounded
  slack (so a coincidental keyword overlap across two unrelated paragraphs
  does not masquerade as a split), AND
- the matched right tokens appear **in order** across `R[a..b]` (an LCS check,
  not just multiset containment — guards against shuffled-word false positives).

Symmetric **merge** detection (N:1, two adjacent left blocks → one right block)
is the reverse and shares the machinery.

### Representation (the new capability)

Add a 1:N split op to the edit script. The minimal, apply-verifiable shape:

```csharp
// One left block expands into an ordered list of right blocks; the FIRST is the
// "host" carrying the surviving Modify token-diff, the rest are split-offs.
internal sealed record IrSplitBlockOp(
    string LeftAnchor,                       // the one left block
    IrNodeList<string> RightAnchors,         // ≥2 right blocks, document order
    IrNodeList<IrTokenDiff> SegmentDiffs);   // per right block: the token diff of
                                             // L's corresponding segment vs that block
```

`SegmentDiffs[i]` is the token diff of the *i-th slice* of `L`'s token stream
(sliced at the in-order LCS boundaries from detection) against `RightAnchors[i]`.
The inserted paragraph mark(s) and any net-new content (the math run in WC-1830)
fall out as Inserts within the appropriate segment diff — exactly the oracle's
`Deleted "When you click…"` + `Inserted "\nA=πr2\nWhen you click…"` shape.

A `MergeBlockOp` is the mirror (N left anchors, one right anchor).

### Ripple implications (why this is a real project, not a patch)

- **Apply-verifier** (`IrEditScript` apply path): applying a split must
  reconstruct N right blocks from one left block — splice L's surface at the
  segment boundaries and interleave the per-segment inserts. New apply branch +
  invariant (`apply(split, [L]) == [R[a..b]]` at text level). This is the load-
  bearing correctness gate; it must be green over the corpus both directions.
- **Markup renderer** (`IrMarkupRenderer`): a split must emit native tracked-
  changes markup that Word/LibreOffice accept⇒right, reject⇒left. The natural
  encoding is `w:del` of the removed paragraph mark region + `w:ins` of the new
  paragraph mark + math, threaded so the prefix/suffix text stays un-marked
  (Equal). This is the subtle part — paragraph-mark revision markup
  (`w:rPr/w:del` on the paragraph mark `w:pPr`) is its own corner case.
- **Revisions renderer** (`IrRevisionRenderer`): project a split op to the
  oracle's del/ins revision pair (the segment inserts/deletes coalesced).
- **JSON** (`IrEditScriptJson`): serialize/round-trip the new op kinds.
- **Fuzzer** (`IrDiffFuzzTests`): add a split/merge mutation class and assert
  the own-oracle apply+JSON invariants plus cross-engine differential parity.

### Cost / determinism

Detection is gap-bounded: for a gap of G free blocks the adjacent-run scan is
O(G²) slices × cached tokenization, the same G²/2 class the in-order refinement
already documents. Deterministic by construction (integer-indexed runs, no
dictionary enumeration for output; the LCS tie-break is the existing patience-
sort discipline).

## Recommendation

Schedule as a **Phase-2 follow-on** ("M2.6 — sub-paragraph split/merge
alignment") once the public surface ships and the engine has burned in as the
opt-in. It is a self-contained capability with a clear apply/markup/JSON/fuzz
contract, and it closes the WC-1450 + WC-1830 deviations at the engine grain (the
only correct level) rather than coarsening Fine-mode output. Until then both
deviations are retained with this sketch referenced; the IR's per-block account
is internally consistent and correct, just coarser at the split boundary than the
oracle's flat atom LCS.
