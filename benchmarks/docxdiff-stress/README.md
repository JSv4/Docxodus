# DocxDiff performance stress harness

A standalone harness that times the `DocxDiff` pipeline against one heavyweight legal
`.docx` and its own generated edit variants, and proves that an optimization changed
nothing about the output. It is deliberately **not** part of `Docxodus.sln`, so it never
affects CI, packaging, or the warning baselines.

Its sibling `benchmarks/complex-form-doc` answers "does the toolchain get this document
right?". This one answers "how fast, and where does the time go?".

## Variants

The generator edits the source package's XML directly rather than driving `DocxSession`,
so the generator's own cost never lands in the measured numbers and the same input always
produces byte-identical variants.

| Case | What it is | What it stresses |
|---|---|---|
| `identical` | no change at all | the byte-identical fast paths |
| `light` | 8 scattered word-level edits | a typical counsel pass |
| `heavy` | an edit in every fifth paragraph | many small in-place modifications |
| `churn` | an edit in every second text node | the token differ |
| `reorder` | 24 blocks relocated across the body | move detection and the LIS spine |
| `structural` | 20 paragraphs deleted, 20 inserted | gap filling and in-gap pairing |
| `footnotes` | an edit in every second footnote paragraph | the note-scope diff |
| `rewrite` | every paragraph's text replaced | the worst case: nothing aligns cheaply |

Plus an N-way `DocxDiff.Consolidate` over four of those variants as reviewers, which is
where read cost scales with reviewer count.

## Output parity

Timing a change is only half the job; the other half is proving the change was free. Each
run records SHA-256 digests of the redline package, the rendered revision list, the
edit-script JSON, and the four consolidate products, for every variant:

```bash
# before the change
dotnet run -c Release --project benchmarks/docxdiff-stress -- doc.docx --baseline before.json
# after
dotnet run -c Release --project benchmarks/docxdiff-stress -- doc.docx --check before.json
```

`--check` exits 2 on any mismatch and prints which product diverged. A perf change that
prints `[parity] OK` produced byte-identical output on every case.

## Corpus differential — the regression evidence

The eight generated variants answer "did this change the diff of one heavyweight legal form".
They do not answer "did this change the diff of anything else", and the reference document
happens to carry **no tracked revisions, no tables and no drawings** — three shapes the
engine's riskiest paths are keyed on. `--corpus` closes that: it runs every document in a
directory through the products and digests the results, so two builds can be compared.

```bash
# baseline on one build
dotnet run -c Release --project benchmarks/docxdiff-stress -- --corpus TestFiles --baseline before.json
# ...switch builds, then
dotnet run -c Release --project benchmarks/docxdiff-stress -- --corpus TestFiles --check before.json
```

> **Building the baseline branch.** The harness reaches internal types, so the branch you take the
> baseline on needs `<InternalsVisibleTo Include="DocxDiffStress" />` in `Docxodus/Docxodus.csproj`.
> Any branch predating that line fails to compile the harness with a wall of `CS0122`, which looks
> like a harness bug and is not one — add the line to the checkout you are baselining, or the
> "before" half of the differential cannot be produced at all.

Against `TestFiles/` each document contributes:

- **four pairwise comparisons** — an edited variant under default settings, the document against
  itself (the byte-identical shortcuts), and the edited variant again under
  `PreAcceptInputRevisions` and `PreserveInputRevisions`, the only settings that reach the
  revision-transform path — times three products (redline, revision list, edit script);
- **one N-way consolidate** of two reviewers off that document as base, times four products
  (redline, consolidated revision list, conflicts, consolidated edit script). The consolidate path
  has its own reader fan-out, merger and markup renderer; not one of the pairwise products touches
  any of them, so without this the whole N-way half of the engine sits outside the differential;
- **one rotated comparison** — the edited variant again with exactly ONE `DocxDiffSettings`
  property moved off its default, chosen by `document index mod N`. The surface is twenty
  properties and a cross-product of them explodes, so each is exercised on roughly `678/N`
  documents for one extra comparison per document rather than twenty;
- **one rotated consolidate** (issue #632) — the same argument applied to the N-way surface:
  the pairwise variations wrapped in a `DocxDiffConsolidateSettings` (which COMPOSES
  `DocxDiffSettings`, so every wrapped variation is reachable on this path too), extended with
  the two conflict-resolution policies that exist only there.

A thrown exception is recorded too, as `FAIL <Type>: <message>`. Malformed and unsupported
fixtures are expected to throw; a change in **which** exception they throw is still a regression.

### The unit is an observation, not a digest

A product call has more observable effects than its return value, and that is not a detail — it is
where the #616 regression lived. That change added a fast path returning an empty revision list for
byte-identical documents, and skipped the compatibility pre-flight on the way. For two identical
documents "no revisions" is the correct answer before and after, so **the return value never moved
and all 8,136 digests agreed, correctly**. The harness was not broken and it was not unlucky: it
answered the question it was built to answer, and the defect was somewhere else.

So the recorded unit is a record with one field per channel, and adding a channel is one field on
one type applied to every product and every document at once:

| Channel | Records | Default |
|---|---|---|
| `Result` | the return value's digest, or `FAIL <Type>: <message>` | — |
| `Warnings` | the compatibility pre-flight's feature ids **for that same call** | `none` |
| `InputMutation` | which side's bytes the call moved | `clean` |
| `OrderVariance` | whether the memoized `DocxDiffComparison` agrees with the static | `stable` |

`Warnings` closes the #616 hole directly: a product that stops running the pre-flight records
`none` where it used to record a report, which no output digest can tell you because the output was
already right. It is captured from the call under observation rather than from a second run of every
product, which is both cheaper and stricter.

`InputMutation` exists because `IrReader.Read` documents *"the caller's `DocumentByteArray` is left
byte-for-byte unchanged"* and `PreAccept` promises *"the input is untouched"*, and nothing verified
either — while the engine moves toward sharing one parsed snapshot across stages, which is exactly
the direction that ends in mutating a caller's bytes.

`OrderVariance` guards a risk #616 created. Each static used to recompute from scratch; now they
delegate to a `DocxDiffComparison` that memoizes one provenance-bearing IR snapshot and shares it
across every product, so `GetRevisions()` then `ToRedline()` traverses different shared state than
the reverse order. Each observation asks a single comparison for all three products in **reverse**
order and requires each to match what the static produced — which also pins
`CreateComparison(l, r, s)` against `DocxDiff.Compare(l, r, s)`, a class corpus mode otherwise never
reaches, since it only ever calls the statics.

Every run prints how many observations record a **non-default** value per channel. A channel that
is never anything but its default across the whole corpus is coverage nobody should count.

### What a differential cannot do

**It validates change, not correctness.** The baseline comes from `main`. A behaviour that was
already wrong on `main` agrees with itself and prints `[parity] OK`. Digests catch drift; only
assertions catch violation — which is why the unit tests matter more than any digest column.

And one class is invisible to every harness: a **leaked resource**. `IrReader.Read` once
constructed its `OpenXmlMemoryStreamDocument` before the `try`, so handing it an `.xlsx` leaked the
open package that the previous `using` disposed. No digest, in any harness, can see that. Nor can
tooling: `CA2000` was measured against the exact leak shape and **does not fire** at default
`AnalysisMode`, at `Recommended`, at `All`, or with `dotnet_diagnostic.CA2000.severity = warning`
set explicitly. Turning on broad CA analysis is separately a non-starter — `AllEnabledByDefault` on
`Docxodus.csproj` produces thousands of warnings, almost entirely legacy noise. That class needs a
targeted test per known throw path, and someone to notice the path exists.

### Confirm the harness can actually fail

An always-green check is worthless, so verify it detects a change before trusting a pass. **One
perturbation is not enough**, because the three products are sensitive to different halves of the
pipeline — a control that moves two of them can leave the third completely unexercised:

| Perturbation | redline | revisions | editscript | `Warnings` |
|---|---|---|---|---|
| drop the `w:t` skip in `UnidHelper.ContentSignature` | **0%** | 82% | 84% | 0% |
| swap the snapshots handed to `IrMarkupRenderer.Render` | **64%** | 0% | 0% | 0% |
| skip the pre-flight on the identical-bytes shortcut | **0%** | **0%** | **0%** | fires |

The same applies to the two newer channels, and both were verified the same way: mutating an input
inside a product call moves `InputMutation` on 200 of 760 observations, and making one comparison
product disagree with its static moves `OrderVariance` on 400 of 600. A channel that has never been
seen to fire is a channel nobody should trust.

The consolidate rows are the cautionary tale (issue #632). They recorded `Warnings` and
`OrderVariance` as `n/a` on two premises that were both false when written down — the consolidate
settings type COMPOSES `DocxDiffSettings`, so `settings.Diff` carries the compatibility
subscription like any other, and since #617 the four statics each delegate to a single-use
`DocxDiffConsolidation` whose caller-held form shares one memoized read/merge across products.
That inert channel is what let the #629 gap — three of the four N-way entry points never running
the pre-flight — hide behind thousands of rows that looked as though the question had been asked
and answered. Both channels are wired now, and the fire drill was repeated for the N-way half:
disabling the consolidate pre-flight moves 1,832 of 15,594 observations — every consolidate row
that records a real warning (916 default-mode + 916 rotated) — where the pre-change harness
recorded zero. 56 of those also move on the RESULT channel: the rotated `throw-on-compat`
consolidate stops throwing, so the rotation catches the same defect a second, independent way,
exactly as it would have caught #629.

The split is not an accident, and it is worth understanding before trusting any column.
Unids feed block anchors, so corrupting a content signature shows up all over the edit script and
the revision list — but the markup renderer strips `PtOpenXml.Unid` on the way out, so the
rendered package is byte-identical and the redline digest never moves. Conversely the hand-off
only affects rendering: the script and revisions were already computed, so they are untouched
while the redline scrambles. And the third row is why the `Warnings` channel exists at all: a
product that stops warning still returns the right answer, so **every output digest stays green**
while a documented behaviour is gone. Run all of them, or you are validating part of the harness.

**It is not omniscient.** Reintroducing the `wp:docPr/@id` stripping-order bug that
`IrHasherTests.Canonicalize_LoneDocPrId_StillStripped` guards produces **zero** corpus
mismatches: real Word documents emit `<wp:docPr id=".." name=".."/>`, so the lone-attribute
shape that bug needs does not occur in this corpus at all. Corpus parity and unit tests cover
different things; neither replaces the other.

## Running

```bash
dotnet run -c Release --project benchmarks/docxdiff-stress -- path/to/document.docx [options]
```

| Option | Effect |
|---|---|
| `--iterations N` | timed iterations per case (default 5) |
| `--warmup N` | untimed warm-up iterations per case (default 2) |
| `--cases a,b,c` | restrict to named pairwise cases (the N-way case always runs) |
| `--stages` | per-stage timings: IR reads, edit-script build, markup render, revision render |
| `--products` | also time `GetRevisions`, `GetEditScriptJson`, and the fused multi-product pass |
| `--probe` | sub-stage attribution inside `IrReader` and `UnidHelper`, plus the pipeline decomposition |
| `--baseline FILE` / `--check FILE` | write / verify output digests |
| `--out DIR` | write the generated variants as `.docx` for inspection |

Build in `Release`. A `Debug` build's numbers are not meaningful.

### Reading the numbers

`Compare` on this document allocates hundreds of megabytes, so whichever case runs first
after a quiet stretch would otherwise absorb the Gen2 collection for everything before it;
the harness forces a collection between cases to keep them comparable. Medians are the
figure to quote — the max column regularly catches a collection and is not the steady
state.

The `--probe` timings warm each stage past the tiered-JIT promotion threshold before
timing any of them. Without that, the first stages in the series report several times
their steady-state cost, because the later ones inherit code they paid to compile.

## Reference document

Written against the NVCA Model Certificate of Incorporation (October 2025): ~51 pages, 234
body paragraphs, 15,360 elements in `word/document.xml`, 97 footnotes (5,461 elements),
16 abstract numbering definitions, 4 sections, 8 headers and 10 footers. Publicly
available from the NVCA; not committed here. Any comparable form document works.

See `FINDINGS.md` for what the harness measured and what it cost.
