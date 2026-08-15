# Legal workflow evaluation corpus

This corpus is a deterministic acceptance suite for realistic legal-document edits. It separates
two questions that are easy to conflate:

1. Can the Docxodus engine execute and verify a known-good edit plan?
2. Given the same instruction and input, did a model produce an equivalent candidate?

The scripted engine baseline always runs first. Model planning is scored only when that baseline
passes, so an engine regression is never reported as a planning failure.

The CLI and focused xUnit suite use the same artifact-producing runner. Every scenario attempt
publishes a fresh score directory, and `run-summary.json` is written in a `finally` path. If an
operation throws, the bundle preserves the original document, the last valid document checkpoint,
the operation log through the failing operation, diagnostic HTML/diff envelopes, and a deterministic
evaluation receipt. A failed rerun replaces the previous score directory instead of leaving stale
artifacts that look current.

## Run it

The fast subset contains six scenarios and is the ordinary pull-request gate:

```bash
dotnet run --project tools/legal-eval/legal-eval.csproj -- \
  --subset fast \
  --artifacts TestResults/legal-eval/fast \
  --report TestResults/legal-eval/fast-report.json
```

The full subset runs all nine scenarios, including the slower review-thread, consolidation, and
pre-existing-review-state cases:

```bash
dotnet run --project tools/legal-eval/legal-eval.csproj -- \
  --subset full \
  --render \
  --artifacts TestResults/legal-eval/full \
  --report TestResults/legal-eval/full-report.json
```

Use `--scenario <id>` to select one case. To evaluate model output, place one file per scenario at
`<candidate-directory>/<scenario-id>.docx` and add `--candidate-dir <candidate-directory>`.
Requesting model evaluation without one of those files makes the run incomplete and returns a
nonzero exit code. The CLI prints the absolute artifact root at startup and the run summary path at
exit; paths stored inside summaries and indexes are portable and relative to that root.

## Scenarios

| Tier | Scenario | Legal workflow |
| --- | --- | --- |
| fast | `defined-term-targeting` | Change one defined term without touching a lower-case decoy |
| fast | `tracked-notice-amendment` | Amend a notice period with native tracked changes |
| fast | `numbered-clause-insertion` | Insert a styled clause into an existing numbering stream |
| fast | `table-economics` | Change one negotiated amount in a styled economics table |
| fast | `preserve-package-structures` | Preserve signatures, fields, sections, and running content |
| fast | `content-control-fill` | Fill a tagged legal-template content control |
| full | `review-thread-note-cross-reference` | Add a comment/reply, footnote, bookmark, and cross-reference |
| full | `compare-consolidate` | Consolidate two attributed reviewer versions against one base |
| full | `preexisting-review-state` | Edit around existing revisions and comments without flattening them |

Every scenario declares a machine-readable instruction, explicit constraints, canonical expected
artifact IDs, scripted baseline operations, change budgets, and at least one deterministic
invariant. `ScenarioLoader` rejects missing or empty invariant lists; visual inspection is never a
substitute for a probe.

## Scoring contract

Each score contains these metric categories:

| Category | What is measured |
| --- | --- |
| `task_completion` | Scenario-specific OOXML/text invariants and availability of every required output |
| `target_precision` | Normalized-package equivalence to the pinned expected DOCX oracle |
| `unintended_change` | Changed OPC parts and distinct `DocxDiff` anchors against the scenario budget |
| `document_validity` | Material `OpenXmlValidator` errors |
| `redline_reversibility` | `redline-reversibility.interim-text-projection`: accept/reject body/header/footer text projection; not the full issue #464 proof |
| `rendering_regression` | `rendering-regression.html-projection`: sanitized Docxodus HTML equivalence; not a page-layout proof |

Metric evaluation is exception-safe. A malformed candidate becomes failed metrics and still keeps
its original bytes and evidence; it does not abort the remaining corpus. Required `expectedOutputs`
are joined to the produced artifact index by canonical ID and add the
`task-completion.required-artifacts` metric. An absent or unavailable required artifact fails the
score. Optional renderer artifacts remain non-fatal.

The canonical scenario output IDs are `candidate-docx`, `semantic-diff`, `after-html`,
`redline-docx`, `candidate-pdf`, and `redline-proof-v1`. The schema and loader reject aliases so a
scenario cannot silently request an artifact the runner does not know how to verify.

## Evidence from every score

Artifacts are retained for passing, failing, and requested-but-incomplete attempts. Model planning
that was not requested produces no score directory. A score attempt writes to:

```text
<artifact-root>/<scenario-id>/engine-baseline/
<artifact-root>/<scenario-id>/model-planning/
```

Each safe completed edit score contains:

- `input.docx`, `candidate.docx`, and scripted `expected.docx`;
- native `redline.docx`;
- `before.html`, `after.html`, and `target.html` previews;
- full input-to-candidate and input-to-target semantic-diff JSON using the public `DocxDiff`
  edit-script schema;
- `metrics.json`, `operation-log.json`, `summary.md`, and a deterministic
  `evaluation-receipt.json`;
- linked `index.html`/`index.md` views whose entries include media type, size, and SHA-256;
- before/candidate/target/redline PDFs, every rendered page, and page-aligned visual diffs when
  LibreOffice, Poppler, and ImageMagick are available;
- `artifact-status.json`, with relative paths, media types, sizes, SHA-256 values, and explicit
  unavailable reasons.

On an execution failure, the same locations hold the original and last valid checkpoint rather
than pretending the edit completed. Safe checkpoints still receive real HTML and, when requested,
PDF/raster previews; unsafe or absent content receives a viewable diagnostic envelope. A failed
semantic diff is likewise a retained JSON diagnostic with `failed` status, never a missing mystery.

LibreOffice runs with a disposable profile outside the published artifact tree. If a renderer is
missing or conversion fails, `artifact-status.json` records that result rather than inventing a
visual. External renderers are used only for trusted scripted-engine documents; untrusted model
candidates retain sanitized HTML and explicit renderer-unavailable records. Preview HTML has a
restrictive CSP and removes active handlers plus external links/resources.

`evaluation-receipt.json` is evidence for this deterministic test run; it is deliberately not the
future delivery-receipt schema. The same explicit-unavailable rule applies to dependent foundation
work that is not yet available:

- package manifest v1: issue #456;
- expanded semantic diff v2: issue #457;
- delivery receipt v1: issue #458;
- embeddable full-surface redline proof v1: issue #464.

Those stable unavailable IDs are intentional extension points. When a foundation lands, its real
provider can replace the corresponding record without changing scenario execution or the scoring
pipeline.

## Fixture provenance

`fixtures/northstar-cloud-services-agreement.docx` is the pinned input oracle for a synthetic legal
services agreement authored for this repository. Every scenario also names a pinned expected DOCX
and SHA-256; the scorer compares a candidate to that independent oracle, never to output generated
in the same run. The adjacent `.fixture.json` remains a readable source recipe and a focused test
regenerates the pinned input byte-for-byte. The package includes styles, numbering, a negotiated
economics table, two sections, signatures, a tagged content control, a bookmark/link, a footnote,
running header/footer content, an existing comment, and an existing revision. Fixed relationship
IDs and deterministic ZIP output keep entry order, timestamps, bytes, and evidence hashes stable.

`provenance.json` records origin, author, date, license, redistribution permission, pinned source
path/hash, and recipe path/hash. Corpus loading verifies all hashes and fails closed when provenance
or redistribution permission is absent.

## Automation and extension points

`.github/workflows/legal-eval.yml` runs the fast corpus and focused xUnit coverage on ordinary CI.
It also reuses the epic #435 MCP smoke workflow to exercise the same public agent-editing path, and
uploads its before/after DOCX plus traces. The full corpus is opt-in through `workflow_dispatch` and
runs weekly on a schedule with document renderers installed and `--render`, producing all-page
before/candidate/target/redline evidence and visual diffs. Fast and full uploads use
`if: always()`, so successful and failed runs are equally inspectable.

The current package guard enforces bounded raw/expanded/XML sizes, ZIP entry counts, compression
ratios, path safety, duplicate names, and DTD-free XML parsing. It is an adapter seam, not a second
permanent security policy: issue #456 remains the source of truth for the shared package safety
manifest/budget implementation. A package that fails this boundary is retained as evidence but is
not sent to downstream OOXML/diff/HTML parsers or external renderers. Likewise, expanded semantic diff v2 (#457), delivery receipt v1
(#458), and the embeddable full-surface redline proof (#464) remain explicitly gated.

New operations belong in `ScriptedBaselineExecutor`; new deterministic package probes belong in
`EvaluationScorer`. Every new scenario must include at least one compatible deterministic invariant.
Artifact integrations belong behind `ArtifactWriter` and must return either a real, hashed artifact
or an explicit unavailable record. Keep model orchestration outside the baseline executor: model
output enters only through `--candidate-dir`.
