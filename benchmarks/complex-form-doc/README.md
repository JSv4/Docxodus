# Complex-form-document benchmark

A standalone harness that runs Docxodus's agent-facing surfaces against one heavyweight
legal .docx and verifies the invariants that matter for redline work. It is deliberately
**not** part of `Docxodus.sln`, so it never affects packaging or the warning baselines —
but CI does compile it (`ci.yml`, "Build out-of-solution tools and benchmarks") so a
library change cannot rot it unnoticed.

## What it measures

Every row below is one `Bench(...)` label the harness prints as `[bench] <stage>`. The
`ComplexFormBenchmarkContractTests` test in `Docxodus.Tests` asserts this table and
`Program.cs` name exactly the same stages, so neither can drift from the other.

| Stage | Surface | Checks |
|---|---|---|
| `html (footnotes+headers rendered)` | `WmlToHtmlConverter` | timing, output size |
| `markdown projection` | `WmlToMarkdownConverter` | timing, anchor count |
| `DocxDiff compatibility probe` | `DocxDiff.InspectCompatibility` | compatibility warning count |
| `no-edit session round-trip` | `DocxSession` open → save with no edits | text-exact, part inventory identical |
| `tracked session: full edit script` | `DocxSession` with tracked changes on | every scripted edit succeeds, native revisions recorded with the configured author, single accept/reject works, no schema findings added |
| `clean session: same edit script untracked` | `DocxSession` with tracked changes off | timing (produces the "revised" side of the redline) |
| `DocxDiff.Compare` | `DocxDiff.Compare` | timing |
| `DocxDiff.GetRevisions` | `DocxDiff.GetRevisions` | revision count |
| `DocxDiff edit script + semantic changes` | `DocxDiff.GetEditScriptJson`, `GetSemanticChangesJson` | payload sizes |
| `DocxDiff round-trip invariants` | `RevisionProcessor` over the redline | **accept-all ≡ revised text**, **reject-all ≡ baseline text**, no schema findings added |
| `redline -> HTML with tracked-change markup` | `WmlToHtmlConverter` with `RenderTrackedChanges` | timing |

"No schema findings added" is measured against the *source document's own* validator
baseline — real Word documents ship with schema findings of their own (the reference
document carries 80), so zero is the wrong yardstick; parity is the invariant.

## Reference document

`TestFiles/NVCA-Model-COI.docx` — the NVCA Model Certificate of Incorporation
(October 2025), committed in this repository. 234 paragraphs, 94 footnotes, 392
bookmarks, 4 sections, 8 headers and 10 footers across 44 package parts. Any comparable
form document can be passed instead; see the edit-script note below.

## Running

```bash
dotnet run --project benchmarks/complex-form-doc -- TestFiles/NVCA-Model-COI.docx [edits.json] [--out DIR]
```

The edit script defaults to `edits/nvca-coi.json`, which encodes a realistic Series A
counsel pass over the NVCA charter (placeholder fills, a liquidation-preference and
dividend-rate change, a notice-period change, deletion of an optional bracketed
provision, an inserted negotiated paragraph, a formatting-only change that crosses a
cross-reference field, and a span-anchored comment). For a different document, supply an
edits JSON with the same shape: the needles are literal substrings located via
`DocxSession.Grep`.

## Exit codes

| Code | Meaning |
|---|---|
| 0 | every `[check]` line passed; the harness prints `ALL CHECKS PASSED` |
| 1 | usage error — no document path was given |
| 2 | at least one `[check] ... FAIL`; the count is printed on the last line |

A stage that throws is caught, reported as
`[bench] <stage>: FAILED after N ms :: <ExceptionType>: <message>`, and counted the same
way a failed check is — so one broken stage still lets the remaining stages run, and the
run still exits 2.
