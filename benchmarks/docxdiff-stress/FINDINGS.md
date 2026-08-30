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

## Filed follow-ups

| | |
|---|---|
| [#617](https://github.com/JSv4/Docxodus/issues/617) | Reusable snapshot so bulk pipelines read a document once — the lever that decides 10/s |
| [#618](https://github.com/JSv4/Docxodus/issues/618) | `UnidHelper` hashes twice per element for elements the IR never reads back |
| [#619](https://github.com/JSv4/Docxodus/issues/619) | `MarkupCompatibilityNormalizer` full-parses `document.xml` on every call to find nothing |

## On the 10 comparisons/second target

Not reached for a cold pairwise redline of a document this size, and worth being precise
about why. At ~350 ms the `light` case is just under 3 per second; 10 per second means 100 ms.

What that 100 ms would have to absorb: ~45 ms of unavoidable package I/O, plus the edit
script, plus assembling and re-deflating an output package. The **data** products are much
closer — `GetEditScriptJson` and `GetRevisions` are 216 ms and, with both documents already
read, about 60 ms of that is the diff itself. **If the question is "10 analyses per second",
that is reachable now with snapshot reuse; if it is "10 redlined .docx packages per second",
the zip round-trip is a hard floor and the answer on this hardware is no.**

Four things remain, in descending value, none of them taken here. The first three are filed
so they are trackable rather than buried in this file:

1. **Reuse reads across comparisons — [#617](https://github.com/JSv4/Docxodus/issues/617).**
   Nothing in the pipeline lets a caller say "I already
   read this document." Every realistic bulk workload — one baseline against many
   counterparties' markups, a version chain, an N-way consolidate — reads the same document
   over and over. A snapshot type carrying the pre-accepted document plus its IR, accepted
   by `CreateComparison`, removes the read from every comparison after the first. With both
   sides pre-read, the `light` case would be roughly 165 ms — 6 per second — and the data
   products would clear 10 per second comfortably. This is a public API addition and rippl-
   es across every transport (see the ripple checklist in `CLAUDE.md`), which is why it was
   left out of a change set whose remit was to not alter the surface.

2. **Assign Unids only where they are read — [#618](https://github.com/JSv4/Docxodus/issues/618).**
   ~51 ms of every read hashes twice per element
   for all 15,360 elements, but the IR reads a Unid back only on blocks, content controls
   and `w:drawing`. A block's Unid derives from its ancestors, never its descendants, so
   pruning the assignment would leave every anchor byte-identical. It was not taken because
   `UnidHelper` is shared with `DocxSession` and the markdown projection, Unids persist into
   saved packages, and "which elements carry an identity" is exactly the kind of contract
   a green test suite can fail to protect.

3. **Stop full-parsing `document.xml` to prove there is nothing to normalize —
   [#619](https://github.com/JSv4/Docxodus/issues/619).** `MarkupCompatibilityNormalizer`
   gates its parse on the part containing the literal `pPr`, which every real
   `word/document.xml` does, so the gate never closes: ~18 ms per side, ~10% of a
   comparison, spent building a DOM that finds nothing. The reference document has zero
   paragraphs carrying the duplicate-`w:pPr` shape the repair targets.

4. **Spend the other two cores.** Pre-accept and the two IR reads are concurrent now, which
   uses about two of four cores; the build and the render are single-threaded. Per-block
   token diffs are independent and the per-scope reads within one document are nearly so.
   Deliberately NOT filed: it is worth perhaps 15-20% against a real risk of introducing
   order-dependence into a pipeline whose determinism is a documented guarantee, it is
   unavailable in the browser at all (see below), and an issue saying "consider more
   threads" is not a unit of work anyone can pick up. It belongs here as a note, not in the
   tracker.

## The concurrency is not available everywhere

`wasm/DocxodusWasm/DocxodusWasm.csproj` does not set `WasmEnableThreads`, so the browser runtime is
single-threaded. There `Task.Run` does not start a thread — it queues the delegate for the one thread
that would then block on the result, and the runtime refuses rather than deadlocking: the join throws
`PlatformNotSupportedException: Cannot wait on monitors on this runtime`. An unguarded fan-out
therefore does not merely fail to be faster, it fails every browser comparison outright — measured,
by forcing the guard open and watching all ten `npm/tests/docx-diff.spec.ts` cases fail with exactly
that exception. `Docxodus.Internal.ParallelWork` compiles the fan-out out under `WASM_BUILD` and
checks `Environment.ProcessorCount` besides.

That matters for how the gain is attributed: **the read sharing is the larger half, and it does not
depend on threads.** With the fan-out forced off (medians of nine, same box):

| Case | `main` | sequential | concurrent |
|---|---|---|---|
| `light` | 821 ms | 575 ms | 351 ms |
| `heavy` | 868 ms | 604 ms | 395 ms |
| `reorder` | 792 ms | 514 ms | 338 ms |
| `structural` | 803 ms | 531 ms | 334 ms |
| `footnotes` | 959 ms | 717 ms | 510 ms |
| `rewrite` | 1279 ms | 1025 ms | 868 ms |

So roughly 1.4-1.6x from reading each document once, and the rest from reading the two of them at
the same time. Treat the split as approximate — the two columns come from different runs and this
box's medians move by 10-15% with load — but the ordering is stable and the mechanism is not in
doubt: the stage decomposition puts the two reads at ~235 ms sequential against ~134 ms concurrent.

All 36 output digests match on both schedules, which is the point: the schedule is not a semantic
choice.

## Notes for anyone re-running this

- **Measure medians, and force a collection between cases.** A comparison here allocates
  hundreds of megabytes; whichever case runs first after a quiet stretch absorbs the Gen2
  collection for everything before it. Before the harness settled the heap between cases,
  the same case reported 334 ms and 1048 ms in consecutive runs.
- **Warm past tiered JIT before timing, not just twice.** An early draft of the sub-stage
  probe reported `UnidHelper` at 108 ms and the package parse at 48 ms; with every stage
  given the same promotion budget those became 51 ms and 13 ms. Stages measured first in a
  series pay for code the later ones inherit already compiled.
- **One document is not a corpus.** The digests over eight generated variants of the reference
  document say nothing about the shapes that document lacks — and it lacks tracked revisions,
  tables and drawings, which between them gate the revision-transform path, the table differ
  and every media import. `--corpus TestFiles` is the answer: 678 documents, 8,136 digests,
  ~4 minutes. Run it before believing a perf change to this engine is free.
- **`--check` is the point.** A perf change to a diff engine is only interesting if the diff
  is unchanged, and "the tests still pass" is a weaker claim than "all 36 output digests are
  byte-identical across eight edit shapes and a four-way consolidate".
- **One negative control validates one half of a pipeline.** The first perturbation used to
  prove the corpus check could fail — corrupting a content signature — moved 82-84% of the
  revision and edit-script digests and **0%** of the redline digests, because Unids are stripped
  from the rendered package. The redline column was entirely unvalidated until a second control
  (swapping the snapshots handed to the renderer) moved 64% of it and none of the others. When a
  check covers several products, perturb each one's inputs before believing any of them.
- **Verify a suspected failure mode; do not infer it from timing.** The threading problem above was
  first blamed on a CI Playwright job that had run 51 minutes without finishing. That reasoning was
  wrong twice over: the job was cancelled by the next push rather than timing out, and the actual
  defect throws immediately rather than hanging. Building the WASM bundle with the guard forced open
  and running one spec settled it in ten minutes and produced the exact exception text now quoted in
  `ParallelWork`. A slow job is evidence of a slow job.
- **Digest parity on one document is not coverage.** One optimization here — skipping the
  attribute sort for elements carrying a single attribute — was briefly wrong: canonicalization
  strips `wp:docPr/@id` by consulting `attribute.Parent`, so deciding an attribute's fate after
  detaching it silently changes the verdict. Every digest still matched, because the reference
  document contains no drawings. The rule was only pinned once a unit test gave `wp:docPr` a
  lone `@id` — the existing one always paired it with `@name`, which took the sorted path. Read
  the code an optimization touches for shapes the corpus does not contain.
