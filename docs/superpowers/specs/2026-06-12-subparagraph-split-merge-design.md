# Sub-Paragraph Split/Merge Alignment — Resolved Design (M2.6 follow-on)

**Date:** 2026-06-12
**Status:** IMPLEMENTED (M2.6, 2026-06-12 — commits 07060d8..HEAD on `feat/diff-m24`; scoreboard
179/179 genuine, deviation catalog empty). The MUST-FIX gate is closed item-by-item in the
**IMPLEMENTATION OUTCOME** section appended at the end; the canonical description of the
as-implemented algorithm (including deltas from §§2–4 below) is
`docs/architecture/ir_diff_engine.md` § "1:N paragraph split / N:1 merge".
Originally: DESIGN-RESOLVED — implementation deferred. This spec resolves the
open op-model / detection / apply / markup / wire questions left by the original
sketch into a single buildable contract, then ran an adversarial DESIGN REVIEW
(separate reviewer agent) whose verified findings are appended at the end. The
review surfaced two holes on the path the *actual fixtures* exercise (the
cell-path apply proof, F3.2; the WC022 identity-reservation interaction, F4.2)
and two oversold claims (N:M-impossibility F1.1, merge-symmetry F2.1) — all
carried as MUST-FIX preconditions on the implementer in the review section
below. The design is buildable; those four items are the gate.
**Companion:** [`2026-06-11-ir-diff-layout-program-plan.md`](../plans/2026-06-11-ir-diff-layout-program-plan.md)
(decision-log entry M2.5-followon), `docs/architecture/comparison_engine.md`,
[`2026-06-12-diff-m26-residuals-and-split-design.md`](../plans/2026-06-12-diff-m26-residuals-and-split-design.md)
(Task 3 — the deliverable this spec satisfies).
**Deviations it closes:** WC-1450, WC-1830 (the two retained GetRevisions
deviations whose root cause is 1:N paragraph split).
**Scope guard:** 1:N split + N:1 merge ONLY. N:M (split-and-merge in one gap) is
explicitly OUT (see §1.5).

---

## Problem (the established evidence)

`WmlComparer` runs a single whole-document atom-level LCS. The IR diff engine
aligns at **block (paragraph) grain** first (`IrBlockAligner`), then token-diffs
*inside* each 1:1-paired block (`IrTokenDiffer`). When one before-paragraph's
content **migrates across two after-paragraphs** (a paragraph SPLIT — the user
pressed Enter mid-paragraph, or inserted a paragraph that swallowed the break),
the oracle's flat atom LCS handles it naturally; the IR's block model **cannot
pair one left block with two right blocks**, so it surfaces an extra
whole-paragraph revision.

### WC-1830 — `WC/WC041-Table-5.docx` → `…-Mod.docx`

A single table cell.

| | LEFT cell blocks | RIGHT cell blocks |
|---|---|---|
| 0 | `Video provides … your point. When you click … to add.` (ONE paragraph) | `Video provides … your point. ` (split prefix) |
| 1 | `You can also type a keyword …` | `` (empty math paragraph holding `A=πr2`) |
| 2 | `` (empty) | `When you click … to add.` (split suffix) |
| 3 | — | `` (empty) |

**Oracle GetRevisions (2):** `Deleted "When you click … to add."` +
`Inserted "\n"` (the inserted paragraph mark — empty/newline text). (The oracle's
delete-side variant; see §3 for the exact produced XML — the oracle deleted the
original collapsed paragraph mark and re-inserted the tail as new paragraphs, but
the shared `Video provides…` prefix run stays **Equal/unmarked**.)

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

Identical mechanism, identical residual (+1). Oracle 7 / IR 8. (This fixture
also contains a *second*, cleaner split cell — the `Second `-prefixed cell — that
the oracle renders with the **anchored-split** strategy; see §3.2. Together the
two fixtures exercise BOTH oracle strategies, which is why they are the design's
target shapes.)

## Why no bounded fix exists within the 1:1 model

The crux is provable, not a tuning question:

1. **`IrEditOp` is strictly 1:1** — one `LeftAnchor`, one `RightAnchor`, one
   `IrTokenDiff` over one left paragraph vs one right paragraph
   (`Docxodus/Ir/Diff/IrEditScript.cs:97`). There is **no shape** that maps one
   left block to two right blocks.
2. The oracle's result requires crediting `When you click` (RIGHT suffix) as
   Equal against LEFT[0]'s SUFFIX while *also* crediting `Video provides`
   (RIGHT prefix) as Equal against LEFT[0]'s PREFIX. That is one paragraph
   token-diffed against TWO — a 1:2 token diff by definition.
3. **Render-time coalescing cannot recover it.** `IrRevisionRenderer`'s
   compat coalescing merges *contiguous same-direction whole-block ins/del runs*
   and trims common affixes *within one ModifyBlock's token diff*. Here the ops
   are interleaved (Insert, Insert, Modify, Delete) and the Modify already binds
   LEFT[0] to the wrong single right block; no regrouping of 1:1 ops
   reconstructs a 1:2 Equal crediting.
4. **Re-pairing the other half is symmetric** — whichever half LEFT[0] binds to,
   the *other* half becomes a whole-paragraph Insert. 1:1 keeps at most one of
   the two; the oracle keeps both.

Therefore matching the oracle is **engine-level 1:N split semantics**, a new
edit-script capability.

---

## 1. Op-model decision

**DECISION: dedicated `IrSplitBlockOp` / `IrMergeBlockOp` as two NEW
`IrEditOpKind` members on the EXISTING `IrEditOp` record, carrying a multi-anchor
list + per-segment diffs in trailing nullable fields. NOT an n-ary
generalization of the existing 1:1 fields. 1:N + N:1 only; N:M out.**

Concretely, extend the record (trailing, nullable, default-null — preserving all
positional constructor call-sites, exactly as `TableDiff`/`TextboxDiffs` were
added at `IrEditScript.cs:104-105`):

```csharp
internal enum IrEditOpKind
{
    // … EqualBlock, FormatOnlyBlock, ModifyBlock, InsertBlock, DeleteBlock,
    //    MoveBlock, MoveModifyBlock …
    SplitBlock,   // one LEFT block → ordered N≥2 RIGHT blocks
    MergeBlock,   // ordered N≥2 LEFT blocks → one RIGHT block
}

internal sealed record IrEditOp(
    IrEditOpKind Kind,
    string? LeftAnchor,                    // SplitBlock: the one left block; MergeBlock: null
    string? RightAnchor,                   // MergeBlock: the one right block; SplitBlock: null
    IrTokenDiff? TokenDiff,
    int? MoveGroupId,
    bool? IsMoveSource,
    IrTableDiff? TableDiff = null,
    IrNodeList<IrTextboxDiff>? TextboxDiffs = null,
    IrNodeList<string>? SplitMergeAnchors = null,   // SplitBlock: the N right anchors,
                                                    //   doc order; MergeBlock: the N left anchors
    IrNodeList<IrSegmentDiff>? SegmentDiffs = null);// one per SplitMergeAnchors entry

/// The token diff of the i-th slice of the single side's token stream against the
/// i-th multi-side block. Mirrors IrTextboxDiff(IrNodeList<IrEditOp>) as the
/// "nested per-child diff" precedent (IrEditScript.cs:125).
internal sealed record IrSegmentDiff(IrTokenDiff Diff);
```

### 1.1 Why dedicated kinds, not n-ary generalization of `IrEditOp`

Evaluated against every surface the op must contract with:

| Criterion | n-ary generalize `IrEditOp` (make `LeftAnchor`/`RightAnchor` plural) | dedicated `SplitBlock`/`MergeBlock` kinds (chosen) |
|---|---|---|
| **Apply-verifier** (`IrEditScriptVerifier.Verify`, `:67-171`) | EVERY existing per-kind branch (Equal/Modify/Insert/Delete/Move) would have to re-handle the now-plural anchor fields; the load-bearing `ReferenceEquals(actual, rightBlock)` count/order assertion (`:133-140`) is written one-right-block-per-op and would need rewriting for all kinds. High blast radius on the correctness gate. | One NEW `case` pushes N reconstructed tuples; all existing branches untouched. The count/order/reference-identity assertion proves `apply(split,[L])==[R[a..b]]` for free (it already iterates per produced right block). |
| **Markup emission** (`IrMarkupRenderer.RenderBlockOp`, `:213-275`) | The dispatch switch is keyed on `Kind`; plural anchors don't change dispatch but every branch must defend against "is this the 1:N case?". | One new dispatch arm; the existing arms are oblivious. |
| **JSON wire** (`IrEditScriptJson`, `:60-93`/`:227-250`) | Pluralizing `leftAnchor`/`rightAnchor` is a BREAKING wire change to the existing op shape — every serialized op gains/loses a field; the deterministic byte-identity contract (fuzzer-enforced) churns for all kinds. | New optional `splitMergeAnchors`/`segmentDiffs` arrays appear ONLY on split/merge ops (the `textboxDiffs` precedent, `:78-91`); existing ops serialize byte-identically. No migration. |
| **Consolidate-forward compat** (downstream consumers iterate `op.Kind`) | A consumer that exhaustively switches on `Kind` keeps compiling but silently mishandles the now-plural fields of kinds it thought were 1:1. | A consumer's `switch` gets a compile-time gap for the new kinds (or hits its `default`) — the change is VISIBLE, not silent. |
| **Revisions-surface mapping** (`IrRevisionRenderer.RenderBlockOp`, `:259-301`) | Same as markup: every arm must re-check arity. | One new arm projecting to the oracle del/ins pair. |
| **M-suffix / N:M scope creep** | Plural-on-both-sides INVITES N:M (the fields already allow it); reviewers will ask "why not N:M?" and the type permits it. | The kinds are named 1:N (`SplitBlock` has exactly one `LeftAnchor`) and N:1 (`MergeBlock` has exactly one `RightAnchor`). N:M is UN-REPRESENTABLE without a third kind — the scope boundary is enforced by the type, not by convention. |

The deciding factor is the **apply-verifier and the wire-determinism contract**:
both are tightened by keeping the 1:1 ops byte-for-byte unchanged and adding the
1:N capability as a strictly additive kind. This mirrors how `Moved` was added
(a new `IrEditOpKind`, not an overload of `InsertBlock`/`DeleteBlock`).

### 1.2 Why the single host carries `LeftAnchor`/`RightAnchor` and the multi-side
       goes in `SplitMergeAnchors`

A `SplitBlock` is fundamentally one-left → many-right. Putting the singular side
in the existing `LeftAnchor` field means `AssertAnchorsResolve`
(`IrEditScriptVerifier.cs:447-456`), `IrRevisionRenderer`, and any anchor-walker
that reads `op.LeftAnchor` still see the left block with zero changes; only code
that needs the N right blocks reaches into `SplitMergeAnchors`. The mirror holds
for `MergeBlock` (singular `RightAnchor`, plural lefts). This keeps the "field
presence by kind" doc table (`IrEditScript.cs:80-96`) a strict superset of
today's.

### 1.3 SegmentDiffs slicing

`SegmentDiffs[i]` is the token diff of the *i-th slice* of the single side's
token stream (sliced at the in-order LCS boundaries computed during detection)
against `SplitMergeAnchors[i]`. The inserted paragraph mark(s) and any net-new
content (the math run in WC-1830) fall out as Inserts within the appropriate
segment diff. An EMPTY connective right block (the bare paragraph-mark carrier a
split inserts) gets a segment diff that is all-Insert over zero content tokens —
which the revision renderer's empty-mark prune (`:266-278`) already suppresses
into a paragraph-mark `\n` revision, matching the oracle.

### 1.4 Why a `MergeBlock` mirror now (N:1), not later

Merge (N adjacent left → one right) is the byte-reverse of split and shares 100%
of the detection machinery (run the containment scan with left/right swapped) and
~100% of apply/markup/JSON. Shipping it in the SAME milestone is cheaper than a
second pass: the fuzzer's `MergeParagraphs` mutation already produces it (the
inverse of `SplitParagraph`), and Word's merge markup is the inverse mark
(delete the joining paragraph mark — accept merges). Excluding it would leave a
known-symmetric gap that the fuzzer would immediately surface.

### 1.5 N:M explicitly OUT

A gap that simultaneously merges some blocks and splits others (N left ↔ M right,
neither side singular) is **out of scope** and UN-REPRESENTABLE in this op model
(neither `SplitBlock` nor `MergeBlock` admits a plural-on-both-sides shape). The
detection algorithm (§2) only fires when exactly ONE side of the candidate run is
a single block. An N:M gap falls through to the existing 1×1-residue /
surplus-insert/delete behavior (the current, coarser-but-correct account). This
is the deliberate scope ceiling: the two retained deviations are both clean 1:N;
N:M has no fixture evidence and no oracle target shape to match.

---

## 2. Detection algorithm

### 2.1 Where: `IrBlockAligner.FillOneGap`, between the similarity pass and the
       1×1-residue rule

`FillOneGap` (`IrBlockAligner.cs:298-400`) runs in order: in-order refine
(Unchanged, then FormatOnly) → drop consumed → **`SimilarityPair`** (greedy
best-score ≥ `BlockSimilarityThreshold`, `:334-335`) → collect leftovers
(`:338-345`) → unambiguous-table residue (`:363-373`) → **1×1-residue rule**
(`:383-392`) → surplus Deleted/Inserted (`:394-399`).

The split/merge containment scan slots in **after `SimilarityPair`
(`:335`) and before the 1×1 rule (`:383`)**, operating on the
`leftoverLeft`/`leftoverRight` lists — the same residue zone as the table-residue
scan. Rationale for the position:

- **After similarity** so a genuinely better 1:1 pairing always wins first (a
  split is only proposed over blocks similarity declined to pair).
- **Before the 1×1 rule** because the 1×1 rule only fires at *exactly* one free
  left + one free right; a genuine 1:2 split has one free left + ≥2 free right, so
  it currently falls THROUGH to surplus (Deleted + 2×Inserted) at `:394-399` —
  which is precisely the +1 deviation. Catching it before surplus is the fix.

### 2.2 The containment criterion

For each unpaired left block `L`, scan **maximal runs of ADJACENT unpaired right
blocks** `R[a..b]` (in right-document order), treating empty/whitespace-only
paragraphs as connective tissue (a split inserts an empty paragraph-mark
carrier). Compute coverage of `L`'s content-token stream by the run:

```
inOrderLcsLen(L, R[a..b]) = | LCS( tokens(L), tokens(R[a]) ++ … ++ tokens(R[b]) ) |
coverage(L, R[a..b])      = inOrderLcsLen / |contentTokens(L)|
foreignSlack(L, R[a..b])  = |run content tokens NOT in the LCS| / |run content tokens|
```

A **split candidate** fires when ALL hold:

1. `b > a` — a genuine 1:N with N≥2 **non-empty** right blocks (empty carriers
   between them do not count toward N but ARE absorbed into the op);
2. `coverage ≥ SplitCoverageThreshold` (proposed **0.90** — `L`'s words are
   almost entirely present, IN ORDER, across the run);
3. `foreignSlack ≤ SplitForeignSlack` (proposed **0.34** — bounded net-new
   content per the WC-1830 math insert; guards against two unrelated paragraphs
   that coincidentally share keywords masquerading as a split);
4. the matched tokens appear **in order** across `R[a..b]` — guaranteed by using
   an LCS (not a multiset bag) for `inOrderLcsLen`, so a shuffled-word false
   positive cannot reach the coverage bar.

**Merge** is the mirror: swap L↔R (one unpaired right block `R`, a maximal run of
adjacent unpaired left blocks `L[a..b]`), same thresholds.

The thresholds MUST be swept over the WC corpus during implementation (the M2.4b
precedent: the low-coverage-coarsening 0.67/≥8 thresholds were swept to a stable
plateau). The 0.90/0.34 values above are the design's starting point, chosen so
both fixtures fire and no currently-passing row is mis-classified; the
implementer treats them as tunable and pins the swept values in a test.

### 2.3 Interaction with the identity-reservation pass (`InOrderRefine` Phase 1)

`InOrderRefine` Phase 1 (`:472-504`, the WC022 crossing fix) pairs same-unid
blocks before first-fit. A split's PREFIX segment keeps the original paragraph's
unid (Word preserves the unid on the first half of a split). So the prefix-right
block may ALREADY be paired (Unchanged or Modified) against `L` by the time the
containment scan runs — leaving only the split-off TAIL as a free Insert. The
detection must therefore handle TWO entry states:

- **(a) Fully free split:** `L` is unpaired and the whole `R[a..b]` run is free
  (WC-1830's delete-and-reinsert cell). The scan proposes a `SplitBlock` directly.
- **(b) Prefix-already-paired split:** `L` was paired Modified with `R[a]`
  (prefix) by identity-reservation/similarity, and `R[a+1..b]` are free Inserts
  whose concatenated content is `L`'s un-matched tail. The scan must **detect
  that the existing `L↔R[a]` Modify plus the adjacent free inserts together form
  a split**, and PROMOTE the Modify entry to a `SplitBlock` absorbing
  `R[a+1..b]`. This is the WC-1450 `Second `-prefix cell (the anchored-split
  shape, §3.2). Concretely: after the scan picks its candidates, a reconciliation
  step checks each just-formed `Modified` entry whose left block's *unmatched tail
  tokens* equal the concatenated content of the immediately-following free Insert
  run, and rewrites the pair into a `SplitBlock`.

State (b) is why detection is a SCAN-AND-RECONCILE, not a pure pre-similarity
pass: the prefix may be claimed by an earlier stage, and the split is only
visible once you consider the Modify entry together with its trailing inserts.

### 2.4 Interaction with similarity pairing order, move detection, low-coverage
       coarsening

- **Similarity pairing order:** unchanged — split detection consumes only what
  similarity declined. A split candidate that would steal a block similarity
  paired is impossible (similarity ran first and won).
- **Move detection (`DetectCrossGapMoves`, `:555+`):** runs AFTER all gaps are
  filled, over global leftovers. A split's blocks are consumed within their gap
  (the `SplitBlock` op marks `L` and `R[a..b]` paired), so they are no longer
  leftovers and cannot be re-claimed as a move. A CROSS-gap split (prefix and
  split-off tail land in different gaps) is OUT — treated as N:M-adjacent and left
  to the existing behavior; no fixture exercises it.
- **The new identity-reservation pass:** see §2.3 — it is the reason for the
  reconcile step.
- **Low-coverage coarsening at render** (`IrRevisionRenderer.IsLowEqualCoverage`,
  floor 0.67 / min 8 content tokens): a `SplitBlock`'s prefix segment Modify has
  HIGH equal coverage (it's the un-split prefix, mostly Equal), so it does NOT
  trip the low-coverage path; the split renders as the oracle's del/ins pair, not
  a coarsened whole-block del+ins. Verified against the target shapes in §4.2.

### 2.5 Determinism and cost bounds

- **Determinism:** integer-indexed runs in right-document order; the LCS uses the
  existing patience-sort discipline; ties broken by smallest `(a,b)` then highest
  coverage (a total order). No dictionary enumeration feeds output.
- **Cost:** for a gap of `G` free blocks the adjacent-run scan is O(G²) candidate
  slices × cached tokenization × O(token²) LCS per slice — the same G²/2 class the
  in-order refinement already documents, bounded by gap size (gaps are small in
  practice; a hard cap `SplitMaxRunLength` ≈ 8 bounds pathological gaps). The
  reconcile step (§2.3b) is O(#Modified entries) with a single tail-token
  comparison each.

---

## 3. Oracle-derived target shapes (the produced markup our renderer must emit)

Run on the two fixtures (`WmlComparerSettings { AuthorForRevisions = "Oracle" }`),
the oracle's produced `word/document.xml` shows the **paragraph-mark revision** is
an **EMPTY `<w:ins .../>` or `<w:del .../>` nested at `w:p / w:pPr / w:rPr`**
(Docxodus additionally stamps `pt14:Status="Inserted"|"Deleted"` on `w:pPr`).
Inserted body text is `w:ins > w:r > w:t`; deleted body text is
`w:del > w:r > w:delText` (`xml:space="preserve"` where affix whitespace matters);
Equal runs and an unchanged mark are left bare (`<w:r>…`, `<w:pPr/>`).

**This is exactly what `IrMarkupRenderer.MarkParagraphMark`
(`IrMarkupRenderer.cs:1071-1099`) already emits** — an empty `w:ins`/`w:del`
inside `pPr/rPr`, RevisionProcessor-recognized (accept of a deleted mark MERGES
the paragraph with the next; reject restores it). The split op reuses this method
verbatim; no new markup primitive is needed.

The oracle uses **two strategies**, and both fixtures supply target XML:

### 3.1 Delete-and-reinsert (WC-1830; WC-1450 delete-side cell)

The oracle deletes the original collapsed paragraph mark (`pPr/rPr/w:del`) and
re-inserts each new paragraph with inserted marks + inserted bodies; the leading
shared sentence run stays Equal/unmarked. Excerpt (WC-1830 cell, abridged):

```xml
<w:p>                                            <!-- the original, mark deleted -->
  <w:pPr pt14:Status="Deleted"><w:rPr><w:del .../></w:rPr></w:pPr>
  <w:r><w:t xml:space="preserve">Video provides … prove your point. </w:t></w:r>  <!-- EQUAL -->
  <w:del …><w:r><w:delText>When you click … to add.</w:delText></w:r></w:del>
</w:p>
<w:p>                                            <!-- inserted empty paragraph -->
  <w:pPr pt14:Status="Inserted"><w:rPr><w:ins .../></w:rPr></w:pPr>
</w:p>
<w:p>                                            <!-- inserted math paragraph -->
  <w:pPr pt14:Status="Inserted"><w:rPr><w:ins .../></w:rPr></w:pPr>
  <w:ins …><m:oMathPara> … A=πr² … </m:oMathPara></w:ins>
</w:p>
<w:p>                                            <!-- inserted text paragraph -->
  <w:pPr pt14:Status="Inserted"><w:rPr><w:ins .../></w:rPr></w:pPr>
  <w:ins …><w:r><w:t>When you click … to add.</w:t></w:r></w:ins>
</w:p>
```

### 3.2 Anchored-split (WC-1450 `Second `-prefix cell — the cleanest 1:N shape)

The first paragraph's mark stays UNMARKED (`<w:pPr/>`, Equal), the shared run
stays Equal, and only the split-off paragraph carries the inserted mark +
inserted body:

```xml
<w:p>                                            <!-- prefix: mark + shared run Equal -->
  <w:pPr/>
  <w:ins …><w:r><w:t xml:space="preserve">Second </w:t></w:r></w:ins>  <!-- leading insert -->
  <w:r><w:t>Video provides … prove your point.</w:t></w:r>            <!-- EQUAL -->
</w:p>
<w:p>                                            <!-- split-off second half -->
  <w:pPr pt14:Status="Inserted"><w:rPr><w:ins .../></w:rPr></w:pPr>
  <w:ins …><w:r><w:t>When you click … to add.</w:t></w:r></w:ins>
</w:p>
```

### 3.3 What our renderer must emit + the paragraph-mark corner case

The renderer should emit the **anchored-split** shape (§3.2) — it is the minimal,
cleanest encoding and the apply/reject semantics are clean: REJECT removes the
inserted split-off mark, re-merging the paragraph (reconstructs LEFT);
ACCEPT keeps the split. The delete-and-reinsert shape (§3.1) produces the SAME
accept/reject document and the same GetRevisions counts, so matching the oracle's
*counts/types/texts* (the scoreboard's assertion) does NOT require byte-matching
which strategy the oracle chose — it requires emitting *a* valid split markup that
round-trips and yields the oracle's revision count. The design therefore commits
to anchored-split and treats §3.1 as an equivalent the markup-parity tests accept.

**The corner case the sketch flagged — paragraph-mark revision markup — is
RESOLVED:** it is the existing `MarkParagraphMark` (empty `w:ins`/`w:del` in
`pPr/rPr`). The split op's only new markup work is to call it on the synthesized
split-off paragraph mark (`RevKind.Ins`) for a split, and on the joining mark
(`RevKind.Del`) for a merge — both already implemented for the Move destination
path (`IrMarkupRenderer.cs:816`). No new OOXML primitive.

---

## 4. Surface contracts

### 4.1 Apply-verifier (`IrEditScriptVerifier.Verify`, `:67-171`)

The apply invariant is: applying the script to LEFT reconstructs RIGHT at text
level, with the reconstructed right-producing ops listing right blocks in
right-document order, proven by `Assert.Equal(actualRight.Count,
reconstructed.Count)` + per-position `ReferenceEquals(actual, rightBlock)`
(`:133-140`).

**Extension:** add a `case SplitBlock` to the `Verify` loop and to
`ReconstructBlocks` (`:390-431`) that pushes **N** reconstructed tuples — one per
`SplitMergeAnchors[i]` — each reconstructed from the SINGLE left block by
replaying `SegmentDiffs[i].Diff` over `L`'s i-th token slice (the prefix slice is
mostly Equal-copy; later slices are Insert-driven for the split-off/new content).
The count/order/reference-identity assertions then prove
`apply(split,[L]) == [R[a..b]]` with NO new assertion code. `MergeBlock` is the
mirror: one reconstructed tuple from N left blocks (concatenate their slices).

**New invariant to add (`AssertSplitMergePairing`, analogous to
`AssertMovePairing` `:463-514`):** a `SplitBlock` has exactly one non-null
`LeftAnchor`, `SplitMergeAnchors.Count ≥ 2`, `SegmentDiffs.Count ==
SplitMergeAnchors.Count`, and every anchor resolves (extend `AssertAnchorsResolve`
`:447-456` — which currently reads only `op.LeftAnchor`/`op.RightAnchor` — to walk
`SplitMergeAnchors`). Mirror for `MergeBlock`.

### 4.2 Revisions-surface mapping (`IrRevisionRenderer.RenderBlockOp`, `:259-301`)

Derived from the oracle's GetRevisions on the fixtures:

- **WC-1830 GetRevisions (2):** `Deleted "When you click … to add."` +
  `Inserted "\n"` (the inserted paragraph mark, empty/newline text).
- **WC-1450 GetRevisions (7):** the relevant split contributes `Inserted
  "When you click … to add.\n"` + the paragraph-mark `Inserted "\n"` /
  `Deleted ""` pairs (the full 7 is the whole-cell account; the split's
  contribution is the del/ins-of-tail + paragraph-mark revisions).

**Mapping:** a `SplitBlock` renders to:
- the prefix segment's token diff (mostly Equal → no revision; any prefix edit →
  its own ins/del),
- for each split-off segment: an `Inserted` revision over that segment's inserted
  body text (via the existing `RenderTokenOps`, `:541-588`),
- a paragraph-mark `Inserted "\n"` revision for each NEW mark — the existing
  empty-mark handling (`:266-278`) already produces the `\n` text the oracle
  reports.

`MergeBlock` mirrors: a `Deleted` over the joined tail + a paragraph-mark
`Deleted ""`/`"\n"`. The compat coalescing (`RenderBlockOpList`, `:148-177`) needs
no change — the split's synthesized inserts are emitted through the same
per-segment path and the empty-mark prune already matches the oracle.

### 4.3 JSON schema addition (`IrEditScriptJson`, `:60-93` / `:227-250`)

`WriteOp` gains, after the `textboxDiffs` block (the precedent, `:78-91`), two
OPTIONAL arrays emitted ONLY for split/merge ops:

```jsonc
{
  "kind": "SplitBlock",
  "leftAnchor": "p:cell…:…",          // singular side
  "splitMergeAnchors": ["p:cell…:r0", "p:cell…:r1", "p:cell…:r2"],
  "segmentDiffs": [ { "diff": { …tokenDiff… } }, { … }, { … } ]
}
```

`ReadOp` (`:227-250`) gains symmetric optional `TryGetProperty("splitMergeAnchors")`
/ `("segmentDiffs")` blocks; the constructor call (`:249`) gains the two new
positional args. The enum name round-trips transparently (written via
`Kind.ToString()` `:63`, read via `Enum.Parse` `:229`). Existing 1:1 ops serialize
**byte-identically** (no field churn) — preserving the fuzzer's
determinism/round-trip contract for every non-split op.

### 4.4 Compat-mode rendering

In `WmlComparerCompatible` mode the split renders to the oracle's del/ins +
paragraph-mark revision set (§4.2). In `Fine` mode the split renders its full
per-segment account (prefix Equal + per-split-off Insert + each new mark) — which
is STRICTLY MORE faithful than today's coarse Deleted + 2×Inserted surplus, and is
the engine-grain-correct account the sketch argued for. The two modes diverge only
in coalescing granularity, exactly as they do for ModifyBlock today.

---

## 5. Risk register, test plan, effort, out-of-scope

### 5.1 Risk register

| # | Risk | Likelihood | Mitigation |
|---|---|---|---|
| R1 | **False-positive split** — two unrelated paragraphs sharing keywords classified as a split, suppressing real insert/delete revisions. | Med | The in-order LCS (not bag) coverage ≥0.90 + foreignSlack ≤0.34 + sweep-pinned thresholds; the scan runs only on similarity-declined residue. Fuzzer cross-engine differential catches under-reporting. |
| R2 | **Reconcile mis-promotion (§2.3b)** — a legit Modify+adjacent-Insert wrongly fused into a split. | Med | Tail-token EXACT-match gate (the Modify left block's unmatched tail must equal the inserted run's concatenated content, not merely overlap). |
| R3 | **Apply-reconstruction regression** — the multi-block-per-op branch breaks the count/order assertion for some corpus doc. | Low | The branch is purely additive; full-corpus `IrEditScriptVerifier` run both directions is the gate (already green over the corpus for 1:1). |
| R4 | **Markup round-trip (accept/reject) wrong** — reject does not re-merge the paragraph. | Low | Reuses `MarkParagraphMark` whose accept/reject merge semantics are already RevisionProcessor-tested; add explicit accept→LEFT / reject→LEFT round-trip assertions on both fixtures. |
| R5 | **Threshold blast radius** — a swept threshold flips a currently-passing scoreboard row. | Med | The genuine-pass ratchet (177) is the guard; the sweep must keep PASS ≥ 177 and only convert WC-1450/1830 from DEVIATION to PASS. |
| R6 | **Cross-gap / N:M leakage** — a multi-gap split mis-handled. | Low | Explicitly OUT (§1.5/§2.4); detection requires one singular side within one gap, so cross-gap simply doesn't fire. |
| R7 | **Note/textbox/cell scope** — `ProjectAlignment` is shared, so split fires in note/cell/textbox scopes too; an unexpected scope interaction. | Med | The shared projection is a FEATURE (split-in-cell is exactly WC-1450/1830); but add scoped fuzzer coverage (split inside a table cell and a footnote) to the comparable battery. |

### 5.2 Fuzzer extensions needed

Add `SplitParagraph` + `MergeParagraphs` to `DiffFuzzer.MutationKind`
(`DiffFuzzer.cs:36-65`) with: enum member + `Describe` case (`:77-89`) + weighted
`PickMutation` entry (`:241-257`) + `Apply` case (`:278-388`, split breaks one
`Para` at a word boundary into two; merge concatenates two adjacent `Para`s).
Decide each mutation's `IsComparableClass` (`:109-110`): a split the IR projects
to a clean del/ins+mark pair the oracle matches is COMPARABLE; if the swept
thresholds leave a residue class the engines frame differently, exclude it there
(as `RelocateParagraph` is). The own-oracle battery (`RunOwnOracle`,
`IrDiffFuzzTests.cs:203-224`) needs NO new wiring — apply-verify + JSON
round-trip + determinism flow automatically once the mutation kind exists.

### 5.3 Test plan

1. **Two fixture regression tests** (WC-1450, WC-1830): GetRevisions count/type/
   text == oracle; scoreboard converts both DEVIATION→PASS (genuine-pass ratchet
   177→179, parity floor unchanged at 179).
2. **Markup parity** (`IrMarkupParityScoreboardTests`): produced doc accept→RIGHT,
   reject→LEFT for both fixtures; OOXML schema-valid; paragraph-mark revision
   present.
3. **Apply-verifier** over the full WC corpus both directions stays green; explicit
   `apply(split,[L]) == [R[a..b]]` unit cases.
4. **JSON** round-trip + determinism for split/merge ops; a golden serialized
   split op.
5. **Aligner unit tests:** detection fires for (a) fully-free and (b)
   prefix-already-paired states; does NOT fire for keyword-coincidence
   non-splits; threshold-sweep test pinning the chosen values.
6. **Fuzzer:** split/merge mutation classes green over the seed battery; scoped
   split-in-cell / split-in-note cases.
7. **N:1 merge** symmetric fixture (constructed — no corpus N:1 fixture exists;
   build one from the WC-1450 cell reversed).

### 5.4 Effort estimate

**~5–7 engineer-days**, decomposable:

| Workstream | Est. | Notes |
|---|---|---|
| Op model + JSON (record fields, enum, Write/Read, field-presence doc) | 0.5d | Additive; textboxDiffs is the template. |
| Detection in `FillOneGap` (containment scan + reconcile + thresholds) | 1.5d | The hardest part: state (b) reconcile + sweep. |
| Apply-verifier extension + corpus green | 1.0d | Multi-block-per-op branch + pairing assert. |
| Revision renderer mapping + compat | 0.75d | Mostly reuse; empty-mark prune already there. |
| Markup renderer (reuse `MarkParagraphMark`) + accept/reject round-trip | 0.75d | Low — primitive already exists. |
| Fuzzer mutation classes + comparability decision | 0.75d | Plus scoped cell/note cases. |
| Threshold sweep + scoreboard ratchet + docs | 0.75d | Pin swept values; CHANGELOG; arch doc update. |

Drivers of variance: the threshold sweep (R5) and the prefix-already-paired
reconcile (R2) are the two places that can balloon if the corpus surfaces an
unanticipated split shape.

### 5.5 Out of scope (explicit)

- **N:M** (split-and-merge in one gap) — §1.5; un-representable by design.
- **Cross-gap splits** (prefix and tail in different alignment gaps) — §2.4.
- **WC-1920** (cross-run word coalescing) — a tokenizer-class deviation, unrelated.
- **Move-and-split** (a split whose halves also relocate) — no fixture; would
  compose `MoveBlock` with `SplitBlock`, deferred.
- **Re-byte-matching the oracle's strategy choice** (§3.1 vs §3.2) — we emit
  anchored-split; parity is on counts/types/texts + round-trip, not strategy.

---

## Recommendation

Schedule as the **M2.6 sub-paragraph split/merge** capability. It is self-contained
with a clear apply/markup/JSON/fuzz contract, closes WC-1450 + WC-1830 at the
engine grain (the only correct level), and the load-bearing markup primitive
(`MarkParagraphMark`) and the empty-mark revision prune already exist — most of the
risk is in detection thresholds, not new machinery. Until implemented, both
deviations remain RETAINED with this spec referenced.

---

## Adversarial DESIGN REVIEW

*(produced by a separate reviewer agent, grounded in the actual code; upgrades the
header to DESIGN-RESOLVED. Findings the implementer MUST close before relying on a
claim are marked. The body §§1–5 above are left as the PROPOSED design; this
section is the adjudication layer over them.)*

### Findings — Axis 1 (op-model soundness)

**F1.1 [major] — "N:M is un-representable by the type" is FALSE.** The added
`SplitMergeAnchors`/`SegmentDiffs` are nullable+additive alongside nullable
`LeftAnchor`/`RightAnchor`; nothing in the *type* forbids a `SplitBlock` that also
sets `RightAnchor`. N:M is rejected only by the builder never emitting it and by
`AssertSplitMergePairing` rejecting it — the same "enforced by convention" property
the §1.1 table assigns to the n-ary option as a weakness. **MUST-FIX:** restate
§1.1/§1.5 as "N:M is rejected by `AssertSplitMergePairing` + never emitted; the
field set physically permits it, so the pairing-assert is load-bearing." Make
`AssertSplitMergePairing` assert `RightAnchor is null` for `SplitBlock` and
`LeftAnchor is null` for `MergeBlock`.

**F1.2 [major] — anchor-walkers silently skip `SplitMergeAnchors`.** `op.LeftAnchor`
walkers see only the singular side; the N real anchors live in `SplitMergeAnchors`
and go un-processed unless each walker is extended (`AssertAnchorsResolve`
`:447-456` reads only Left/Right). §1.2 frames "walkers keep working" as a benefit;
for a resolution/invalidation walker it means "silently skips N anchors."
**MUST-FIX:** §1.2 must enumerate EVERY anchor-walker and mark each
extended-or-provably-anchor-free; until then this is an open under-processing risk.

**F1.3 [minor] — `IrSegmentDiff(IrTokenDiff Diff)` is a scalar wrapper with no
second field;** `IrNodeList<IrTokenDiff>?` carries the same data. The
`IrTextboxDiff` analogy is inapt (that wraps a *list* + is a named scope). Either
drop it or name the concrete future field that justifies it.

### Findings — Axis 2 (scope creep)

**F2.1 [major] — N:1 merge is scope creep dressed as symmetry.** Every evidentiary
deviation (WC-1450/1830) is a SPLIT; no corpus deviation closes merge, and §5.3-7
admits the only merge test is a hand-authored fixture (validating an implementation
against a fixture authored to match it is unfalsifiable). §1.4's "100%/~100%
shared" is overstated — the §2.3b reconcile is split-only. **MUST-FIX:** either
gate merge on a real corpus fixture surfacing (fast-follow), or re-justify it
explicitly as apply-verifier *confidence* (exercising the N→1 reconstruction path
split's mirror depends on) and own that framing — not as deviation-closure.

**F2.2 [major] — the §2.3b reconcile can drift toward N:M via overlapping runs.**
The "exactly one singular side" ceiling is asserted but not enforced locally:
nothing stops two adjacent reconciled splits sharing a right block. **MUST-FIX:**
§2.3b must state the guard preventing overlap, `AssertSplitMergePairing` must verify
no right anchor appears in two ops' `SplitMergeAnchors`, and a two-adjacent-splits
fuzzer case must prove the ceiling.

### Findings — Axis 3 (apply-verifier semantics)

**F3.1 [blocker] — "proves apply==[R[a..b]] for free, NO new assertion code" is
false.** The body `Verify` switch (`:67-128`) has no `default`/exhaustiveness, so an
unhandled `SplitBlock` makes the count assert at `:133` FAIL (a new `case` is
mandatory). The order/identity assert (`:140` `ReferenceEquals`) holds only if the
op's N tuples land at the correct *interleave* positions in the flat
`right.Body.Blocks` — a NEW builder-ordering obligation, not free. **MUST-FIX:**
§4.1 must state the builder ordering contract (a `SplitBlock`'s N right blocks must
be right-contiguous at the op's position) + add the pairing assert; downgrade the
headline to "no new text-equality assert, but a new pairing assert + a
builder-ordering obligation."

**F3.2 [major, path-critical] — the spec cites the wrong reconstruction function;
the fixtures are CELL-path, which has no identity proof.** `ReconstructBlocks`
(`:390-431`) is the cell-internal path (flat string list, text-equality only — NO
`ReferenceEquals`); the body switch (`:67-128`) is the one with the identity proof.
WC-1450/1830 are CELL splits (§3.1/§3.2), so the headline "ReferenceEquals proves
apply for free" **does not apply on the path the actual fixtures take.**
**MUST-FIX:** §4.1 must handle BOTH switches, acknowledge the cell path proves only
text-equality, and decide whether cell-split needs a strengthened identity check.

**F3.3 [major, closeable] — slice boundaries are computed at detection but NOT
serialized;** apply must recover them. Closeable IFF each `SegmentDiff[i].Diff` is a
*complete* token diff whose Delete+Equal ops partition L's i-th slice exactly (no
gaps/overlaps), making boundaries implicit in the diff ops. **MUST-FIX:** §1.3/§4.1
must mandate this partition invariant, and `AssertSplitMergePairing` must check that
the concatenation of all segments' left-side (Delete+Equal) tokens equals L's full
token stream exactly once, in order.

### Findings — Axis 4 (regression surface)

**F4.1 [major] — the threshold sweep is asserted feasible, not shown.** 0.34
foreignSlack is tuned to ONE data point (WC-1830's math insert); no evidence a
single (coverage, slack) pair fires on both fixtures without flipping any of the 177
genuine-pass rows that also flow through `FillOneGap`. **MUST-FIX:** demote 0.90/0.34
to "starting hypothesis; the sweep is a GATE that may prove no stable plateau exists
(then the milestone is blocked or thresholds go scope-specific)." Require the sweep
to report the margin to the nearest flip.

**F4.2 [major, reject-order regression risk] — §2.3b under-models the WC022
identity-reservation fix (commit 697611e).** Phase 1 pairs a byte-identical prefix as
**Unchanged/FormatOnly**, NOT Modified — but §2.3b's reconcile checks *Modified*
entries only, so the cleanest fixture shape (§3.2 anchored-split, prefix run
"stays Equal") is MISSED, leaving the tail stranded = the +1 persists. Worse,
promoting an identity-reserved pair into a split changes that paragraph's REJECT
reconstruction (must now re-merge, not restore-in-place), risking re-introducing the
reject-order instability WC022 just closed. **MUST-FIX:** §2.3b must cover the
prefix-paired-as-Unchanged/FormatOnly state, preserve the WC022 reject-order
invariant, and add a both-direction round-trip assert mirroring the WC022 regression
test. Add **R8** to the register: "reconcile promotes an identity-reserved (WC022)
pair → reject-order regression."

**F4.3 [minor, verify-first] — the empty-mark prune is BODY-scope-only
(`IrRevisionRenderer.cs:320`, WC-1190), but the fixtures are CELL splits.** §1.3
assumes the prune suppresses the empty connective mark "matching the oracle"; if
cells follow the note-scope "do NOT prune" rule, a cell split's empty connective
mark yields an EXTRA revision vs the oracle's `\n`. **MUST verify** whether the prune
fires in cell scope before relying on "no new markup work."

### Claims the reviewer VERIFIED SOUND

- **`MarkParagraphMark` reuse (§3.3):** `private static`, emits empty `w:ins`/`w:del`
  in `pPr/rPr` at correct schema order, idempotent (`:1097`), already invoked on the
  move-destination path. **Nit:** that path uses `RevKind.MoveTo` (`:816`), not
  `RevKind.Ins`; the split needs `RevKind.Ins` — a code path move does NOT exercise,
  so "already implemented" overstates reuse by one grade. Fix the §3.3 quote.
- **JSON additive-only wire (§4.3):** confirmed `WriteOp` emits every field
  conditionally; existing ops stay byte-identical; `textboxDiffs` is a true optional-
  array precedent. Sound.
- **Detection insertion point (§2.1):** confirmed order Similarity→leftover→table-
  residue→1×1→surplus; a 1-left+≥2-right split genuinely falls through 1×1 into
  surplus, so slotting before the 1×1 rule is the evidence-backed fix point.
- **Why 1:1 can't fix it (§"Why no bounded fix"):** `IrEditOp` is strictly singular;
  the 1:2 token-diff impossibility argument is sound.

### Adjudication

The design is BUILDABLE. The gate before implementation is the six MUST-FIX items
(F1.1, F1.2, F2.1/F2.2, F3.1, F3.2, F3.3, F4.1, F4.2) and the one verify-first
(F4.3). The two genuinely load-bearing ones are **F3.2** (the apply proof the spec
headlines does not cover the cell path the fixtures use — the spec must either
strengthen the cell-path verifier or stop claiming the identity proof there) and
**F4.2** (the WC022 interaction can both miss the fix and regress reject-order). Neither
invalidates the op-model decision (§1) or the detection placement (§2.1), both of
which the reviewer independently verified sound.

---

## IMPLEMENTATION OUTCOME (M2.6, 2026-06-12)

Shipped on `feat/diff-m24` (op model + JSON `07060d8`, alignment kinds + pairing assert `70ff633`,
segmenter `8f36bfa`, detection `1a7bea9`, projection + verifier `436272a`, revision rendering
`adfcd71`, markup `129dea4` (+`eb4b8da` refactor), default-on + sweep `ea3369f`, fuzzer + interleave
fix `5867082`). GetRevisions scoreboard **179/179 genuine** (WC-1450/WC-1830 closed; the catalog is
EMPTY and `board.Deviation == 0` is asserted); markup round-trip allowlist down to the single
oracle-crashes fixture; full suite 2000/0/1; 1000-seed own-oracle fuzz green; Release build clean.

### MUST-FIX adjudication closure

- **F1.1 (N:M "un-representable" overstated)** — CLOSED as the review demanded: §1.1/§1.5's claim is
  restated here — N:M is *physically representable* by the nullable fields and is rejected by
  `IrEditScriptVerifier.AssertSplitMergePairing` (a `SplitBlock` must carry a null `RightAnchor`, a
  `MergeBlock` a null `LeftAnchor`; no anchor may appear in two ops' `SplitMergeAnchors`) plus the
  builder never emitting it. The pairing assert is the load-bearing scope ceiling and is invoked by
  `Verify` on every corpus/fuzz case.
- **F1.2 (anchor-walker enumeration)** — CLOSED: the audit table lives as a comment above
  `AssertSplitMergePairing` (`IrEditScriptVerifier.cs`), every grep-surfaced walker
  extended-or-proven-anchor-free; `AssertAnchorsResolve` walks `SplitMergeAnchors` (right store for
  splits, left for merges).
- **F1.3 (`IrSegmentDiff` scalar wrapper)** — ADOPTED: no wrapper record exists; `SegmentDiffs` is
  `IrNodeList<IrTokenDiff>?` directly.
- **F2.1 (merge framing)** — OWNED as the review required: `MergeBlock` shipped in the same
  milestone as apply-path CONFIDENCE for the N↔1 reconstruction machinery plus fuzzer coverage
  (`MergeParagraphs` mutation, synthetic fixtures) — NOT as deviation closure; no corpus deviation
  demanded it, and the op-model docs say so.
- **F2.2 (overlap → N:M drift)** — CLOSED: fired members' match slots are stamped immediately and a
  window never admits a consumed index; `AssertSplitMergePairing` rejects any anchor appearing in
  two groups; the two-adjacent-splits case runs at alignment, full-pipeline (Build→Verify→JSON),
  and detection-unit grain.
- **F3.1 (apply "for free" oversold)** — CLOSED per the downgraded claim: a real new `Verify` case
  pushes one reconstructed tuple PER member, and the existing count/order/`ReferenceEquals` loop
  then proves the builder-ordering obligation (the N rights right-contiguous at the op position).
- **F3.2 (cell path identity proof)** — CLOSED: `ReconstructBlocks` (the path the two fixtures
  actually take) gained the split/merge cases AND a produced-right-anchor SEQUENCE assertion
  (right-producing ops must name the right blocks in right-document order) — asserted corpus-wide,
  not just on the split fixtures.
- **F3.3 (slice boundaries not serialized)** — CLOSED exactly as the review's closeable-iff
  condition required: every segment diff is COMPLETE over (slice, member) — re-diffed by the
  ordinary Myers differ per slice — so boundaries are implicit (slice length = Σ non-Insert left
  lengths; merge mirror = Σ non-Delete right lengths); the verifier asserts the slices tile the
  singular stream exactly, and `IrTokenDiffAsserts.AssertInvariants` runs per segment.
- **F4.1 (threshold sweep is a gate)** — CLOSED: `IrSplitThresholdSweepTests` sweeps coverage
  {0.80..0.95} × slack {0.20..0.50} over the WC003 row set; the shipped (0.90, 0.34) sits at the
  grid maximum (104 count-exact rows) on a broad plateau — every ±1-step neighbor attains the same
  count (only slack=0.50 flips a row anywhere in the grid). The plateau + margin assertion runs on
  every test execution; defaults are pinned by a second test.
- **F4.2 (WC022 identity-reservation interaction)** — CLOSED by construction: Unchanged/FormatOnly
  pairs are NEVER promotion candidates (content-equal ⇒ zero unmatched tail ⇒ a trailing insert is
  genuinely new), only a same-gap Modified pairing can be promoted. Regression tests: the
  detection-unit Unchanged-pair negative, plus WC022 both-direction markup round-trips with
  detection ON. (Risk R8, as the review required, is hereby recorded: "reconcile promotes an
  identity-reserved pair → reject-order regression" — mitigated by candidate exclusion, see above.)
- **F4.3 (empty-mark prune scope in cells)** — VERIFIED before reliance: body-table cell paragraphs
  anchor as `p:body:…` (only `p:fn:`/`p:en:` scopes are excluded from the prune), so the prune DOES
  fire in cell scope; pinned by `Cell_scope_empty_mark_prune_fires`.

### Deltas from the proposed design (§§2–4) worth knowing

1. **§3.2's "anchored-split" example cell is NOT a split.** The WC-1450 `Second `-prefix cell's
   before-paragraph never contained the tail (`Score` member-match probe `[11, 0]`): the oracle's
   `Inserted "Second "` + `Inserted "When you click…"` is the ordinary Modify + InsertBlock account,
   which the R2 edge trim correctly preserves. Both true splits in the corpus are the
   suffix-paired state-(b) shape.
2. **Detection is one unified scan, not scan-plus-reconcile.** A candidate may be free OR
   Modified-paired; a qualifying window simply overwrites the pairing. The §2.3(b) "tail-token
   exact-match gate" was superseded by the zero-matched-content EDGE TRIM + the coverage/slack
   gates, which handle both prefix- and suffix-paired states and the interior net-new member with
   one mechanism (R2 is mitigated by the trim: an unrelated edge insert has zero matched content
   and is excluded, shrinking the window below the 2-member floor).
3. **Mark placement:** inserted marks go on paragraphs 0..N−2 and the LAST paragraph keeps the
   original pilcrow (reject's mark-removal merges each paragraph into the NEXT, reconstructing
   LEFT). This differs from the §3.1/§3.2 oracle excerpts' internal strategy but satisfies the same
   accept ≡ right / reject ≡ left contract §3.3/§5.5 adjudicate on. `MarkParagraphMark` is invoked
   with `RevKind.Ins`/`RevKind.Del` — confirming the reviewer's nit that the move path
   (`RevKind.MoveTo`) had not exercised these grades.
4. **Compat-mode account:** the oracle's split contribution is segment-0 inline edits + exactly ONE
   coalesced inserted region (its re-deleted tail coalesces into an adjacent deleted paragraph's
   region when one exists, adding no count) — so the compat renderer emits one coalesced
   `Inserted "\n" + Σ(memberText + "\n")` per split (mirror `Deleted` per merge) rather than
   per-segment + per-mark revisions; Fine mode keeps the per-segment account plus one `"\n"` mark
   revision per added/removed pilcrow (§4.4's account, which also keeps a clean split visible).
5. **An O(1) content-count prefilter** (window content within
   `[coverage·singular, singular/(1−slack)]`) guards the G²-class cost bound — required by the
   adversarial 200×200 fixture once detection went default-on; not anticipated by §2.5.
6. **Fuzzer comparability:** split/merge mutations are EXCLUDED from the cross-engine differential
   class (the engines frame a clean split differently by construction — artifact evidence in the
   500-seed run); §5.2 anticipated exactly this fork. The own-oracle battery covers them on every
   seed. The reshuffled seed stream also exposed a PRE-EXISTING reject-order bug (deletions
   anchored to moved-away lefts restored at the move destination), fixed alongside Task 9.
