# WmlComparer parity baseline — the last differential run

These four files are the verbatim output of the parity scoreboards on their **final run against a
live `WmlComparer`**, captured at commit `59a057f` immediately before the legacy engine was removed
in v11.0.0. They are a historical record, not a live test: the engine that produced the `old=`
column no longer exists in the tree, so nothing can re-derive them.

They are kept because they are the only evidence of *how* `DocxDiff` came to be trusted. When
someone later asks "was the IR engine ever actually checked against the thing it replaced, and on
what", the answer is these numbers and the corpus they were measured over.

| Report | Corpus | Result at freeze |
|---|---|---|
| `IrParityScoreboardTests.txt` | 179 runnable-now WmlComparer cases | 177 PASS, 2 documented deviations, 0 FAIL |
| `IrMarkupParityScoreboardTests.txt` | 39 markup-blocked cases | 39 PASS, 0 deviation, 0 FAIL |
| `ConsolidateParityScoreboardTests.txt` | 84 legacy `Consolidate` cases (WC001 + WC002) | 84 reproduce-PASS, 0 deviation, 0 FAIL |
| `IrVsWmlComparerTests.txt` | 92 WC pairs × 2 directions = 184 comparisons | 117 MATCH, 18 GRANULARITY, 47 DIVERGENT (catalogued), 2 OLD_ERROR |

The two documented deviations and the 47 divergent rows are not defects — each is a catalogued,
explained difference where the IR engine deliberately supersedes the legacy grain (see the
`DIVERGENT sub-buckets` section of the differential report, and `wml_comparer_gaps.md`). The two
`OLD_ERROR` rows are pairs the **legacy** engine threw on and the IR engine handled.

## What replaced them

The differential harness could not survive its own oracle. What survives is
`Docxodus.Tests/Ir/Diff/DocxDiffCorpusBaselineTests.cs` plus its committed
`DocxDiffCorpusBaseline.tsv`: the same 92-pair corpus, both directions, both granularities, with
`DocxDiff`'s own per-kind revision multisets frozen as literals. That test still fails a build on a
regression — it just pins the engine's output rather than comparing two engines. The numbers in the
`.tsv` are the numbers these reports blessed.

Independent oracles that do not depend on `WmlComparer` remain: `tools/diffharness/` (LibreOffice
round-trip verification) and the `eval/` scenario corpus.
