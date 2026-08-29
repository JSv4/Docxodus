# Findings: where DocxDiff's time goes on a heavyweight legal document

Measured on the NVCA Model Certificate of Incorporation (October 2025) — 147 KB packed,
574 KB of `word/document.xml` holding 15,360 elements across 234 body paragraphs, plus 97
footnotes in a 227 KB part (5,461 elements), 16 abstract numbering definitions, 4 sections,
8 headers and 10 footers. Release build, .NET 10.0.11, 4 cores, server GC, single
container. Medians of 9 timed iterations after 4 warm-ups, one process per column.

Timings are indicative of *this* machine; the ratios and the stage attribution are the
durable part.

## What changed

Every number below is from one before/after run of the same harness against the same
document, with `--check` confirming all 36 output digests identical across the two.

| Case | Before | After | |
|---|---|---|---|
| `light` — 8 scattered word edits | 821 ms | **351 ms** | 2.3× |
| `heavy` — every fifth paragraph | 868 ms | **395 ms** | 2.2× |
| `reorder` — 24 blocks relocated | 792 ms | **338 ms** | 2.3× |
| `structural` — 20 deleted, 20 inserted | 803 ms | **334 ms** | 2.4× |
| `footnotes` — half the footnote paragraphs | 959 ms | **510 ms** | 1.9× |
| `churn` — every second text node | 1062 ms | **611 ms** | 1.7× |
| `rewrite` — every paragraph replaced | 1279 ms | **868 ms** | 1.5× |
| `GetRevisions` (light) | 395 ms | **216 ms** | 1.8× |
| `GetEditScriptJson` (light) | 389 ms | **216 ms** | 1.8× |
| all three products fused (light) | 852 ms | **387 ms** | 2.2× |
| `GetRevisions` on identical packages | 371 ms | **0 ms** | — |
| `Consolidate`, 4 reviewers | 1870 ms | **675 ms** | 2.8× |

Allocation per `light` comparison: 528 MB → 276 MB. Per 4-reviewer consolidate: 1337 MB →
710 MB.

The gain tracks how much of a comparison is reading rather than diffing: `structural`, where
the alignment does little work, is 2.4×; `rewrite`, where every block is modified and the
token differ dominates, is 1.5×.

## The finding that mattered: the same document was read four times

A two-way `Compare` spent about 72% of its wall clock inside `IrReader`, reading each of the
two documents **twice**, and each of those reads opened and parsed the package **twice**:

1. `DocxDiffComparison` read both sides with `RetainSources` off to build the edit script.
2. `IrMarkupRenderer` re-read the same two documents with `RetainSources` on, to get the
   source `w:p`/`w:tbl` elements it clones from.
3. Inside every read, deciding the revision view (`Accept`) opened a whole second package,
   parsed every story to scan for revision markup, then threw that parse away and reopened
   the document for the walk.

None of it was needed. `RetainSources` only controls whether `IrProvenance` pins the source
`XElement`, and `IrProvenance` is equality-neutral by construction, so the two snapshots are
node-for-node value-equal — the renderer can be handed the snapshot the script was built
over. And `GetXDocument` caches per part, so scanning the package the walk is about to use
costs nothing extra.

The N-way path had the same duplication multiplied by reviewer count: an N-reviewer
`Consolidate` read `2*(N+1)` packages to compare `N+1`.

## Where the remaining time goes

Stage decomposition of the `light` case after the change (~330 ms), replicating
`DocxDiffComparison`'s own steps:

| Stage | Wall | Share |
|---|---|---|
| IR read, both sides, concurrent | ~134 ms | 41% |
| markup render (excluding the reads it no longer does) | ~113 ms | 35% |
| edit-script build | ~51 ms | 16% |
| pre-accept normalizers, both sides, concurrent | ~25 ms | 8% |

Inside one ~131 ms IR read:

| | |
|---|---|
| `UnidHelper.AssignToAllElementsDeterministic` (main part) | ~51 ms |
| IR walk, hashing, list-marker resolution, note/header scopes | ~60 ms |
| package open + parse of every story the reader consumes | ~20 ms |

Inside the markup render, sampled: 45% is `Buffer.MemmoveInternal` — the package clone and
the deflate round-trip — and 13% is XLinq subtree enumeration. There is no algorithmic
hotspot left there; it is I/O over a 574 KB part.

Inside the edit-script build, sampled: 42% is `IrTokenDiffer.CharWeightedLcs`. That is the
diff actually diffing.

Two floors, for calibration: parsing every part the reader consumes costs ~13 ms, and
opening a package and re-serializing it unchanged costs ~17 ms. So roughly 45 ms of a
330 ms comparison is package I/O that no design avoids.

## On the 10 comparisons/second target

Not reached for a cold pairwise redline of a document this size, and worth being precise
about why. At ~350 ms the `light` case is just under 3 per second; 10 per second means 100 ms.

What that 100 ms would have to absorb: ~45 ms of unavoidable package I/O, plus the edit
script, plus assembling and re-deflating an output package. The **data** products are much
closer — `GetEditScriptJson` and `GetRevisions` are 216 ms and, with both documents already
read, about 60 ms of that is the diff itself. **If the question is "10 analyses per second",
that is reachable now with snapshot reuse; if it is "10 redlined .docx packages per second",
the zip round-trip is a hard floor and the answer on this hardware is no.**

Three levers remain, in descending value, none of them taken here:

1. **Reuse reads across comparisons.** Nothing in the pipeline lets a caller say "I already
   read this document." Every realistic bulk workload — one baseline against many
   counterparties' markups, a version chain, an N-way consolidate — reads the same document
   over and over. A snapshot type carrying the pre-accepted document plus its IR, accepted
   by `CreateComparison`, removes the read from every comparison after the first. With both
   sides pre-read, the `light` case would be roughly 165 ms — 6 per second — and the data
   products would clear 10 per second comfortably. This is a public API addition and rippl-
   es across every transport (see the ripple checklist in `CLAUDE.md`), which is why it was
   left out of a change set whose remit was to not alter the surface.

2. **Assign Unids only where they are read.** ~51 ms of every read hashes twice per element
   for all 15,360 elements, but the IR reads a Unid back only on blocks, content controls
   and `w:drawing`. A block's Unid derives from its ancestors, never its descendants, so
   pruning the assignment would leave every anchor byte-identical. It was not taken because
   `UnidHelper` is shared with `DocxSession` and the markdown projection, Unids persist into
   saved packages, and "which elements carry an identity" is exactly the kind of contract
   a green test suite can fail to protect.

3. **Spend the other two cores.** Pre-accept and the two IR reads are concurrent now, which
   uses about two of four cores; the build and the render are single-threaded. Per-block
   token diffs are independent and the per-scope reads within one document are nearly so.
   Worth perhaps 15-20% and a real risk of introducing order-dependence into a pipeline
   whose determinism is a documented guarantee.

## Notes for anyone re-running this

- **Measure medians, and force a collection between cases.** A comparison here allocates
  hundreds of megabytes; whichever case runs first after a quiet stretch absorbs the Gen2
  collection for everything before it. Before the harness settled the heap between cases,
  the same case reported 334 ms and 1048 ms in consecutive runs.
- **Warm past tiered JIT before timing, not just twice.** An early draft of the sub-stage
  probe reported `UnidHelper` at 108 ms and the package parse at 48 ms; with every stage
  given the same promotion budget those became 51 ms and 13 ms. Stages measured first in a
  series pay for code the later ones inherit already compiled.
- **`--check` is the point.** A perf change to a diff engine is only interesting if the diff
  is unchanged, and "the tests still pass" is a weaker claim than "all 36 output digests are
  byte-identical across eight edit shapes and a four-way consolidate".
- **Digest parity on one document is not coverage.** One optimization here — skipping the
  attribute sort for elements carrying a single attribute — was briefly wrong: canonicalization
  strips `wp:docPr/@id` by consulting `attribute.Parent`, so deciding an attribute's fate after
  detaching it silently changes the verdict. Every digest still matched, because the reference
  document contains no drawings. The rule was only pinned once a unit test gave `wp:docPr` a
  lone `@id` — the existing one always paired it with `@name`, which took the sorted path. Read
  the code an optimization touches for shapes the corpus does not contain.
