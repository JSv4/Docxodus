# Complex-form-document benchmark

A standalone harness that runs Docxodus's agent-facing surfaces against one heavyweight
legal .docx and verifies the invariants that matter for redline work. It is deliberately
**not** part of `Docxodus.sln`, so it never affects CI, packaging, or the warning
baselines.

## What it measures

| Stage | Surface | Checks |
|---|---|---|
| Projections | `WmlToHtmlConverter`, `WmlToMarkdownConverter`, `DocxDiff.InspectCompatibility` | timing, anchor count, compatibility warnings |
| Round-trip | `DocxSession` open → save with no edits | text-exact, part inventory identical |
| Tracked session | `DocxSession` with `TrackedChangeMode.RenderInline` | every scripted edit succeeds, native revisions recorded with the configured author, single accept/reject works, no schema findings added |
| Redline | `DocxDiff.Compare` + revision list + edit-script JSON + semantic changeset | **accept-all ≡ revised text**, **reject-all ≡ baseline text**, no schema findings added |
| Rendering | redline → HTML with tracked-change markup | timing |
| Legacy engine | `WmlComparer.Compare` | same round-trip invariants, tracked as a known-gap regression signal (see FINDINGS.md) |

"No schema findings added" is measured against the *source document's own* validator
baseline — real Word documents ship with schema findings of their own (the reference
document carries 80), so zero is the wrong yardstick; parity is the invariant.

## Reference document class

The harness was written against the NVCA Model Certificate of Incorporation
(October 2025): ~51 pages, 234 paragraphs, 94 footnotes, 392 bookmarks, 201
cross-reference field instructions, 16 abstract numbering definitions, 4 sections,
8 headers and 10 footers. The document is publicly available from the NVCA and is not
committed here — pass any comparable form document on the command line.

## Running

```bash
dotnet run --project benchmarks/complex-form-doc -- path/to/document.docx [edits.json] [--out DIR]
```

The edit script defaults to `edits/nvca-coi.json`, which encodes a realistic Series A
counsel pass over the NVCA charter (placeholder fills, a liquidation-preference and
dividend-rate change, a notice-period change, deletion of an optional bracketed
provision, an inserted negotiated paragraph, a formatting-only change that crosses a
cross-reference field, and a span-anchored comment). For a different document, supply an
edits JSON with the same shape: the needles are literal substrings located via
`DocxSession.Grep`.

Exit code 0 means every check passed; 2 means at least one `[check] ... FAIL` line —
including the two legacy-engine failures that are currently expected on footnote-heavy
documents (FINDINGS.md).
