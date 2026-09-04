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

### 2. Selection — settled by decode (#699), and the answer is that the gate is already right

This point was open when the note was first written; #699 decoded it against the reference compare
corpus (803 base/next/reference triples), and the finding reversed the expected conclusion.

**Retention decays with region size; there is no cliff.** For every region *our* engine replaces
whole — a maximal run of purely one-sided paragraphs with content on both sides — the measurement
asks whether the reference output retained any of that region's shared content words in place,
inside the region's own span:

| our region size | regions sharing a content word | reference retained in place |
|---|---|---|
| 1–2 | 12 | 55% |
| 3–4 | 18 | 50% |
| 5–8 | 29 | 41% |
| 9–16 | 26 | 32% |
| 17–32 | 36 | 28% |
| 33+ | 20 | **10%** |

Word retains less and less as the region grows, and by 33+ members it has essentially stopped. That
is a gradient, not a boundary — which is the first reason no threshold is the criterion.

**Coverage is not the discriminator, and it points the wrong way.** In the reference output's *own*
zero-pair regions (398 of them, where Word paired nothing), content-word coverage between the
deleted and inserted sides *rises* with size — median 0.00 and p90 0.10 at 1–2 members, median 0.15
and p90 0.83 at 33+, with individual regions at coverage 1.0 (a 34×10 region sharing 55 content
words, replaced whole). Word leaves large, heavily overlapping regions entirely unpaired. Any
coverage-based rule would stream exactly those, and be wrong.

**Both literals were then A/B-tested against the corpus, and both are already correct:**

| Literal | Change tested | Result |
|---|---|---|
| `leftParas.Count > 8 \|\| rightParas.Count > 8` | raise to 32 | 8 of 428 regions change arrangement: **3 wins, 3 losses**, 2 with no shared words |
| `crossUnitMatches >= 2 \|\| constructPairs > 0` | relax to `>= 1` | 16 regions start streaming: **0 wins, 1 loss**; 15 of the 16 matched only on words the content filter excludes — precisely the "a lone function word is positional scaffolding" case the gate's own comment describes |

So the size cap is **barely load-bearing** — moving it four-fold is a coin flip on 803 documents, which
means there is no evidence to move it and moving it would churn the corpus baseline for nothing. And
the `>= 2` unit-match floor is **empirically right**: relaxing it buys nothing and costs something.

**The decision.** The criterion is structural, and it already is: the count-equal boundary construct
(a pilcrow pairs when the same number of matched tokens precedes it on each side — no threshold, just
a stream property) is what selects a region, and `>= 2` plus the size cap are outer bounds around it,
not the rule. Point 2 is therefore settled by keeping the gate as it stands, with the two literals
now carrying the measurement that justifies them rather than an admission that the evidence was
missing. What remains genuinely undecided is nothing in this design; what remains *unknown* is
whether a different structural construct would recover the 28–41% of mid-size regions where the
reference retains and we do not — and that is a question about finding a new construct, not about
tuning a gate.

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

- **#699 — Decode reference compare output for large zero-pair regions.** Done; its result is
  point 2 above.
- **#700 — Replace the token-stream mode's tuned gate with the decoded selection criterion.** Closed
  by the same decode: there is nothing to replace. Both literals were A/B-tested and neither is a
  mis-placed threshold.

#694 closed on this note; #699 and #700 closed on the decode that filled in its one open point.
