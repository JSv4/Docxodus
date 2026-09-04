# Token-Stream Arrangement of Regions Without Block Correspondence

Issue #694. This note settles the architecture of `DocxDiff`'s **second arrangement mode** — the
one that renders a region as a single word-level token stream rather than block-wise — and records
which single question is still open, and what would answer it.

It is deliberately short. The mode is not hypothetical: `Ir/Diff/IrCrossParagraphSegmenter.cs`
already implements it, and its class remarks are the specification of the algorithm. This note does
not restate them. It records the decisions *around* the algorithm — placement, selection,
reversibility, provenance, boundaries, cost, moves — and cites the code for the rest.

## The starting point is not a blank page

`IrCrossParagraphSegmenter.SegmentRegion` is already the **general** region form. Its members are a
region's paragraphs in document order per side, and `pairs` lists the aligner's word-matched pairs
as `(left index, right index)`; every member not covered by a pair is one-sided at **any** position
— leading, interior or trailing. Zero pairs is a legal input.

`IrEditScriptBuilder.TryBuildStoryFinalMixedRegionOp` is the caller, and it already admits interior
zero-pair regions. So the shape #694 describes — "one interleaved region in which surviving words
stay in place and paragraph marks land wherever the diff puts them" — is a thing the engine emits
today, for some regions.

What is narrow is not the mechanism. It is **the gate**.

## The seven points

### 1. Placement — settled: a per-region second mode

The token stream is a **renderer-side arrangement of a region the block aligner has already
delimited**, not a replacement for block alignment.

Block-first alignment stays the model of correspondence. Where a block correspondence exists, it is
the right answer and the cheap one, and every guarantee built on it (in-place order monotonicity,
move markup, split/merge groups, table and note recursion) keeps holding untouched. The stream mode
runs only on a region the aligner could not pair, and its output replaces exactly that contiguous
range of entries with one `CrossParagraphRunBlock` op.

A wholesale replacement was considered and rejected. It would have to re-derive, in token-stream
terms, everything the block path already gets right — and the evidence base for the block path
(`DocxDiffCorpusBaselineTests`, the frozen parity reports under
`docs/architecture/wmlcomparer_parity_baseline/`) is not transferable to it. A per-region mode keeps
every existing guarantee intact at the price of needing a selection rule; that price is point 2, and
it is worth paying.

### 2. Selection — **open**, and this is the whole of what #694 leaves undecided

Two literals in the current gate are the tuned thresholds the issue rules out, and they are the only
part of this design that is not settled:

| Where | Literal | What it does |
|---|---|---|
| `IrEditScriptBuilder.TryBuildStoryFinalMixedRegionOp` | `leftParas.Count > 8 \|\| rightParas.Count > 8` | large zero-pair regions never stream |
| `IrCrossParagraphSegmenter.SegmentRegion` | `crossUnitMatches >= 2 \|\| constructPairs > 0` | how much shared text a zero-pair region needs to stream at all |

The size cap is honest about why it exists, in the code comment that carries it: *the decoded stream
constructs all live in small regions, and the replace-gap grammar corpus was validated on exactly
the large ones.* That is a statement about where the evidence is, not about where the boundary
belongs.

**The decision is that the criterion cannot be chosen yet, and that it must be decoded rather than
designed.** Every candidate that can be written from here fails one of the issue's own constraints:

- A **coverage fraction** (matched content characters over the smaller side) is exactly the tuned
  threshold point 2 forbids, and the segmenter has already been burned twice by density and
  matched-char floors — see the pass-1 remark: *"the retain-in-place and yield-to-cross classes
  overlap on coverage, so no threshold separates them."*
- **"Stream every zero-pair region"** — correspondence is absent by construction, so let the stream
  own it — is the cleanest rule available and may well be right, but it directly contradicts the
  only evidence in hand, which says large zero-pair regions arrange with the replace-gap grammar.
- **"Emit what a whole-story LCS gives and accept the interleave"** assumes Word is token-first all
  the way down at every scale. That is the issue's premise, and it is unverified precisely where it
  matters: the decode that produced this engine's rules covered small regions.

The prerequisite is a **decode of reference compare output for large zero-pair regions**: pairs with
little shared text and many paragraphs on each side — a template against a filled-in copy, a memo
against the contract that replaced it, two drafts rewritten rather than edited. Scored on one
question: does the reference output retain surviving fragments in place across paragraph boundaries,
or does it replace the region whole? That answer picks the rule, and it may well be structural
rather than numeric (the count-equal boundary construct below is what a *structural* rule looks
like: no threshold, just "the same number of matched tokens precedes each pilcrow").

Whatever the rule turns out to be, it must be a function of the region's own content — no
document-level state, no per-document switch — so that a region's arrangement does not depend on
what else is in the document.

### 3. Reversibility — settled, and independent of the selection rule

This is the constraint #694 calls the hardest, so the argument is written out rather than asserted.
It matters that it holds **structurally**: widening the gate cannot break `accept ≡ right` /
`reject ≡ left`, because the property does not depend on which regions are selected or on how good
the matching is.

The segmenter emits a list of `IrCrossParagraphCell`s. Three construction invariants carry the
proof, and the walk **asserts** each, returning `null` — caller falls back to the block path — the
moment one is violated:

1. **Single-member cells.** Each cell references at most one left paragraph and at most one right
   paragraph. Cells are closed by boundary events, and every paragraph boundary on either side is a
   boundary event, so a cell can never span one.
2. **Contiguous slices.** Within its member, a cell's slice is a contiguous ascending token range.
3. **Total, disjoint cover.** The expansion walks anchors in stream order and emits every token of
   each side exactly once, on its own side. Per side, therefore, the concatenation of the cells'
   slices in cell order is the side's flat token stream in ascending order — which is, by
   construction, that side's member paragraphs concatenated in document order.

From those three, both directions follow:

- **ACCEPT** keeps every `¶INS` mark, drops every `¶DEL` mark (fusing each `Del` cell into the
  following paragraph), keeps right slices and drops left. By (3) the surviving token sequence is
  the right members' tokens in order. The surviving marks are exactly the right side's boundaries —
  each right boundary is emitted either as a paired boundary or as a `¶INS` — in order. Tokens in
  order plus boundaries in order is the right members' paragraphs, in their document order.
- **REJECT** is the mirror: keep `¶DEL` marks, drop `¶INS`, keep left slices, drop right; by the
  same argument the result is the left members.

Order is preserved, not merely content, because (3) is an ordering statement. That is what #288
requires and what `EnforceInPlaceOrderMonotonicity` enforces at block level; the stream mode needs
no separate order enforcement, because its cells are emitted in stream order and it replaces one
contiguous entry range, leaving every entry outside the region where it was.

Each cell's left/right slices are then re-diffed slice-relative by the ordinary `IrTokenDiffer`, so
the shared markup renderer produces the same `w:ins`/`w:del`/`w:rPrChange` shapes it produces
everywhere else. The stream mode introduces no new markup, and therefore no new
accept/reject semantics to verify.

### 4. Formatting provenance — settled by the per-cell re-diff

A materialised output paragraph carries its own side's properties, and where a cell has both sides,
the per-cell re-diff through `IrTokenDiffer` produces `w:rPrChange` exactly as a within-pair token
diff does. A paragraph mark that survives (a paired boundary) is a retained mark and takes
`w:pPrChange` when the two sides' `w:pPr` differ, the same rule the block path applies to a
`Modified` pair; a `¶INS` mark carries the right side's `w:pPr`, a `¶DEL` mark the left's.

There is no case of "a paragraph that exists in neither input in that exact form" needing an
invented provenance: every cell's content comes from at most one paragraph per side, so there is
always exactly one left and one right candidate, or one of them is empty.

### 5. Structure boundaries — settled, by exclusion

The stream is per-story and flows around structure rather than through it. Membership requires
`IrCrossParagraphSegmenter.IsStreamable` — every inline an `IrTextRun`, no inline section transition
— plus the caller's `HasStructuralCarrier` digest gate. Hyperlinks, fields, note references, images,
opaque inlines, textboxes, tabs, breaks and inline SDT envelopes are all excluded, because they are
zero-width or atomic and a slice boundary falling inside one could double-emit or drop it.

A region containing any of them declines to the block path in full — the collector's maximality test
makes the stream own whole regions and never a fragment of one, so a mid-region table cannot split a
region into two streamed halves. Section breaks are the single transparent exception: they are
story-end metadata, collected through and re-emitted as their own ops after the fused op.

Tables, notes, headers and footers, comment ranges and bookmarks therefore keep their own
reconciliation untouched — the stream never sees them.

### 6. Cost — bounded today, and the point a wider gate stresses first

`LcsCellCap = 1_000_000` DP cells per window; a larger window bails to the block path. Combined with
the `≤ 8` member cap, cost is not currently a live concern.

That changes the moment point 2 is settled in a way that admits large regions. A global word-level
diff over a long story is quadratic in the worst case, and the block aligner's own scale guard
exists for the same reason. Any widening must come with an explicit bound and a fallback to the
block path when it is exceeded — and, because the fallback changes the output, the bound has to be a
function of the region rather than of the machine.

### 7. Moves and split/merge — settled: they stay with the block path

Both are expressed by the aligner, before the renderer arranges anything. A region carrying staged
move sources declines outright today (the region op would displace the flat path's source/deletion
interleave), and split/merge groups are aligner-level 1:N pairings that reach the stream as ordinary
`pairs` entries.

This is deliberate, not a gap: native move markup is a *correspondence* claim about two blocks, and
the stream mode exists precisely for regions where no block correspondence was found. A region that
has a move in it is not that region. If a future decode shows Word emitting move markup inside a
streamed region, that is a new decision, not an extension of this one.

## Test plan

Anything that widens the gate must pass, unchanged and without re-pinning:

- **`DocxDiffFuzzRoundTripTests`** — the generative byte-level `accept ≡ right` / `reject ≡ left`
  round trip, including the paragraph-order assertion from #288. This is the guard the reversibility
  argument above is claiming; if the argument is right, a wider gate cannot move it. Run the wide
  sweep (`DOCXODUS_FUZZ_SEEDS=2000`), not the 250-seed default — the order-sensitive seeds are rare.
- **`DocxDiffCorpusBaselineTests`** — the frozen per-kind revision multiset over the 92-pair corpus,
  both directions, both granularities. A wider gate *will* move entries here, and that is the point:
  every moved entry is a region whose arrangement changed, and each one has to be justified against
  reference output rather than re-pinned because it looked plausible.
- **`DocxDiffGapArrangementTests`** — the arrangement pins for the replace-gap grammar. A region that
  starts streaming is a region that stopped using that grammar, so a moved pin here is the clearest
  statement of what a change actually did.

Plus the two the stream mode already owns: `DocxDiffCrossParagraphTests` and the fusion-ON corpus
battery.

## Child issues

- **#699 — Decode reference compare output for large zero-pair regions.** The prerequisite for
  point 2. Produces the evidence, not a rule.
- **#700 — Replace the token-stream mode's tuned gate with the decoded selection criterion.**
  Blocked on #699. Must land with the cost bound from point 6.

#694 closes on this note. It does not close on the children.
