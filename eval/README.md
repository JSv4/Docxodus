# Workflow evaluation suite

Measures whether an automated document workflow completes the task it was given, leaves
everything else alone, and produces a valid reviewable artifact. Issue
[#466](https://github.com/JSv4/Docxodus/issues/466).

This directory is the **corpus and the contract**. The runner lives in
`Docxodus.Tests/Eval/` and runs as part of the ordinary .NET suite.

## What a scenario is

A scenario is one JSON file under `scenarios/`. [`scenario.schema.json`](scenario.schema.json)
documents the contract; the suite's `EV002` test enforces its substance directly — a non-empty
completion and preservation list, at least one precision bound, and no invariant key outside the
vocabulary the checkers actually read, so a typo cannot silently switch a check off:

| Part | Answers |
|------|---------|
| `fixture` | What document did we start from? |
| `intent` | What would a human have asked for? |
| `steps` | What tool calls perform the task? |
| `invariants` | What makes this run a pass — and what makes it a *precise* pass? |

`intent` is deliberately unused by the scripted caller. It is the prompt an agent run would be
given; the caller executes `steps` verbatim instead. **That split is the point of the suite:** a
failure under the scripted caller is the engine's, and can never be blamed on model planning.
Agent-scored runs reuse the same fixtures and invariants and are scored against this baseline.

## Why the steps are MCP tool calls

`steps` name MCP tools (`docxodus_edit`, `docxodus_create`, …), not the .NET API. The engine
baseline has to be measured where an agent actually meets it, so a gap that only exists in the
tool surface — a missing argument, a wrong precondition, an action that silently no-ops — shows
up here rather than being papered over by a typed call no agent can make.

Steps address blocks by *content*, through `target`, rather than by threading the previous step's
response. A step is therefore independent of the shape of what came before, and a fixture that
drifts fails loudly: if a target matches fewer blocks than the requested `index`, the run errors
out instead of quietly editing a different block and then reporting good precision.

## Fixtures

Fixtures under `fixtures/` are **build scripts in the same step format**, not `.docx` files.
`EvalHarness.BuildFixture` replays one over a blank document. Two consequences:

- The corpus carries no third-party document bytes, so there is no redistribution question to
  answer. See [PROVENANCE.md](PROVENANCE.md).
- A fixture is reviewable as a diff. `EV003` asserts each one builds twice to the same content and
  the same package shape, because a corpus that drifts between runs cannot anchor a baseline.
  Reproducible here means the same *document*, not the same bytes or the same package digest:
  anchor ids and revision-save ids are minted per build, so two honest builds of one script differ
  at the digest layer while being the same document.

## Metrics

Every scenario is scored on the metrics #466 asks for, each sourced from an existing engine
contract rather than a bespoke comparison:

| Invariant | Question | Source |
|-----------|----------|--------|
| `taskCompletion` | Did the intended change land? | `docxodus_search` over every story |
| `targetPrecision` | Did it change *only* that? | distinct anchors in the #457 semantic change set |
| `collateral` | Was unrelated package state preserved? | #456 package manifests, before and after |
| `validity` | Is the deliverable structurally sound? | #463 deliverable gate |
| `reversibility` | Does the redline accept and reject cleanly? | #464 proof |
| `rendering` | Does it still render? | HTML projection |

### Text assertions ask the document, not a rendering of it

`taskCompletion` and `collateral.textPreserved` resolve each needle through `docxodus_search`,
counted over every story, rather than substring-matching the markdown projection. The projection
renders a table *structurally*, so cell text is not present in it as literal prose — an early
version of this suite scored a correct fee-table edit as a total loss for exactly that reason. The
search path is also the one the steps themselves use to find targets, so a scenario asserts
against the same view of the document it edits.

### Reversibility is asserted in layers

`reversibility` always requires that both proof paths complete, that resolving only the
generated revisions leaves pre-existing review state intact (the fixture carries one live
tracked change by another reviewer precisely so this assertion has something to lose), and that
rejecting only the generated revisions is *semantically equivalent to the opening package*
(`rejectMustRestoreBaseline`). Full package equivalence (`mustSucceed`) is **opt-in and
currently off**.

The reason is honesty about what is being scored. The proof needs an intended final, and for a
session-authored redline that document has to be *derived* — here by
`RevisionProcessor.AcceptRevisions`. Requiring full equivalence would therefore assert that the
derivation and the proof's own selective-accept path agree byte-for-byte at the normalized layer,
which is a statement about two engine paths rather than about whether the redline is reversible.
The engine's own `RP001` expects `Success == false` on a generated redline for the same family of
reasons. The reject path has no such excuse — its expected document is the opening package
itself, stated up front — which is why its semantic restoration is asserted by default rather
than being part of the opt-in. Turning `mustSucceed` on, with a fixture whose intended final is
stated rather than derived, is follow-up work.

`targetPrecision` and `collateral` are what make this an evaluation rather than a test. Any edit
can be made to land; the interesting question is what else moved. Each scenario's
`collateral.textPreserved` lists the near-miss content a careless edit would also have caught —
for `term-replacement`, the two *other* occurrences of the defined term.

## Failure artifacts

A failing scenario writes `opening.docx`, `delivered.docx`, `delivered.html`, `delivered.txt`,
`semantic-changes.json`, `verification.json`, and `reversibility-proof.json` (when one was
produced) so a failure is diagnosable without re-running it. Set `DOCXODUS_EVAL_ARTIFACTS` to choose the directory;
otherwise they land under the system temp directory.

## Adding a scenario

1. Write the scenario JSON. Reuse a fixture, or add a build script under `fixtures/`.
2. Declare its invariants. `EV002` requires `taskCompletion`, `targetPrecision`, and a non-empty
   `collateral.textPreserved`: a scenario with nothing scoreable is one that cannot fail, and
   "it looked right" is not an acceptance criterion.
3. Run `dotnet test --filter "FullyQualifiedName~WorkflowEvalTests"`.

## Scope

This is the deterministic fast subset: no PDF, no browser, no network, so it runs on every push.
Deliberately **not** here, and tracked separately:

- Generated-PDF and visual-regression scoring — that is #443's ratchet, and depending on it before
  it exists would bake unstable numbers into this suite's baselines.
- Agent-scored runs and model planning quality, which need this scripted baseline to exist first.
- The larger opt-in corpus, and the remaining #466 scenarios: clause insertion with numbering,
  comment/footnote/cross-reference authoring, content-control templates, N-way consolidation, and
  documents carrying pre-existing revisions.
