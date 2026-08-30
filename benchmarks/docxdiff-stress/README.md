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
- **two pre-flight digests per pairwise product** (edited and identical), recording the
  compatibility warnings each product reports. Output digests cannot see this: a product that stops
  running the pre-flight and a product that runs it and finds nothing produce identical output. The
  identical-bytes shortcuts are where that gap is most tempting, so they are digested too.

That is **678 documents → 22,374 digests**, roughly 15 minutes on four threads.

A thrown exception is digested too, as `FAIL <Type>: <message>`. Malformed and unsupported
fixtures are expected to throw; a change in **which** exception they throw is still a regression.

### The redline digest is rename-invariant, and why

Media and diagram parts imported into a redline are named `P` + a fresh GUID, and the
relationships pointing at them get ids of `R` + a fresh GUID. **`DocxDiff.Compare` is therefore
not byte-deterministic on any document whose redline imports media**, contradicting
`DocxDiffSettings.Deterministic`. This is pre-existing — `origin/main` disagrees with itself on
exactly those documents — and it affects 54 of the 678 fixtures here. Digesting raw package bytes
reports 161 differences on every run and would drown a real regression in noise.

So the redline is digested with those generated names folded to a placeholder, in entry names and
inside XML content, with entries hashed in canonical-name order. Content still has to match
exactly: this hides the naming churn and nothing else. See #621.

### Confirm the harness can actually fail

An always-green check is worthless, so verify it detects a change before trusting a pass. **One
perturbation is not enough**, because the three products are sensitive to different halves of the
pipeline — a control that moves two of them can leave the third completely unexercised:

| Perturbation | redline | revisions | editscript | preflight |
|---|---|---|---|---|
| drop the `w:t` skip in `UnidHelper.ContentSignature` | **0%** | 82% | 84% | 0% |
| swap the snapshots handed to `IrMarkupRenderer.Render` | **64%** | 0% | 0% | 0% |
| skip the pre-flight on the identical-bytes shortcut | **0%** | **0%** | **0%** | fires |

The split is not an accident, and it is worth understanding before trusting any column.
Unids feed block anchors, so corrupting a content signature shows up all over the edit script and
the revision list — but the markup renderer strips `PtOpenXml.Unid` on the way out, so the
rendered package is byte-identical and the redline digest never moves. Conversely the hand-off
only affects rendering: the script and revisions were already computed, so they are untouched
while the redline scrambles. And the third row is why the pre-flight column exists at all: a
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
