# Workflow evaluation suite

Measures whether an automated document workflow completes the task it was given, leaves
everything else alone, and produces a valid reviewable artifact. Issue
[#466](https://github.com/JSv4/Docxodus/issues/466).

This directory is the **corpus and the contract**. The runner lives in
`Docxodus.Tests/Eval/` and runs as part of the ordinary .NET suite.

## What a scenario is

A scenario is one JSON file under `scenarios/`. [`scenario.schema.json`](scenario.schema.json)
is the contract, and it is enforced, not decorative: `EV004` validates every scenario file
against it, `EV005` pins its invariant vocabulary to the checkers' vocabulary in both
directions, and `EV002` enforces its substance directly — a non-empty completion and
preservation list, at least one precision bound, no empty invariant group, and no invariant key
outside the vocabulary the checkers actually read, so a typo cannot silently switch a check off:

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

A tool that takes several anchors — a range format, a move, a bookmark span — declares `targets`
instead: a map whose keys name the arguments and whose values are resolved exactly like `target`.
The key names the argument, so a `targets` entry that also declares `as` fails the scenario as
contradictory.

## Fixtures

Fixtures under `fixtures/` are **build scripts in the same step format**, not `.docx` files.
`EvalHarness.BuildFixture` replays one over a blank document. A fixture may instead declare a
named programmatic `builder` (C# in `Docxodus.Tests/Eval/EvalFixtureBuilders.cs`) for the one
thing the step format cannot express — the tool surface fills content controls but cannot create
one — under the same determinism gate and still without committing bytes. Two consequences:

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
| `collateral` | Was unrelated package state preserved? | #456 package manifests, before and after; `changedPartsMustBeWithin` additionally pins *which* parts the change set may touch |
| `validity` | Is the deliverable structurally sound? | #463 deliverable gate |
| `trackedRevisions` | Is the expected live markup present — including another reviewer's? | `docxodus_track_changes list` on the delivered session |
| `comments` | Are the expected comments, replies, and resolved flags present? | `docxodus_comment list` on the delivered session |
| `redline` | Does the redline a `docxodus_compare` step wrote carry the expected attributed revisions? | `docxodus_track_changes list` over a fresh session opened on `fromPath` |
| `reversibility` | Does the redline accept and reject cleanly? | #464 proof |
| `rendering` | Does it still render? | HTML projection |

`trackedRevisions` and `comments` are read through the same MCP dispatcher the steps drive, so
they assert the agent-surface view of the document, and only what those tools actually report —
replies are entries carrying `parentAnchorId`, resolved state is the tool's own `resolved` flag.
`changedPartsMustBeWithin` names part URIs exactly as the change set reports them
(`/word/document.xml`); an entry outside the list fails loudly, so a misspelled URI cannot pass.

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

## Scorecards and failure artifacts

Every run — pass or fail — writes a per-scenario `scorecard.json`: the metrics that were
actually measured (changed anchors and parts, part-inventory deltas, validity decision, text
match counts, reversibility outcomes). The scorecard is the machine-readable **engine
baseline**: a later agent-scored run of the same scenario — same fixture, same invariants,
planned by a model instead of scripted — is compared against it. Scripted first, agent second,
in that order, is what criterion-level "separate engine correctness from planning quality"
means operationally.

A failing scenario additionally writes `opening.docx`, `delivered.docx`, `delivered.html`,
`delivered.txt`, `semantic-changes.json`, `verification.json`, `revisions.json`,
`comments.json`, and `reversibility-proof.json` (when one was produced) so a failure is
diagnosable without re-running it. Set `DOCXODUS_EVAL_ARTIFACTS` to choose the directory;
otherwise they land under the system temp directory. CI sets it and uploads the directory as a
build artifact when the suite fails; the weekly corpus workflow uploads it always, scorecards
included.

A delivery change receipt is deliberately **not** among the artifacts. A receipt's transaction
lineage must come from authoritative mutation evidence captured at execution time
(`DeliveryTransactionContribution.FromMutationBatchResult` over a recorded batch); the harness
executes steps as independent tool calls and holds only before/after snapshots, and the receipt
contract explicitly forbids reconstructing history from a before/after comparison. Receipt
production belongs to the #465 delivery operation, which owns that evidence.

## Adding a scenario

1. Write the scenario JSON. Reuse a fixture, or add a build script under `fixtures/`.
2. Declare its invariants. `EV002` requires `taskCompletion`, `targetPrecision`, and a non-empty
   `collateral.textPreserved`: a scenario with nothing scoreable is one that cannot fail, and
   "it looked right" is not an acceptance criterion.
3. Run `dotnet test --filter "FullyQualifiedName~WorkflowEvalTests"`.

## The two tiers

Scenarios directly under `scenarios/` are the **fast deterministic subset**: no PDF, no browser,
no network, executed unfiltered on every push. Scenarios under `scenarios/corpus/` are the
**opt-in corpus tier**: executed only when `DOCXODUS_RUN_EVAL_CORPUS=1`, which the weekly
`eval-corpus` workflow (also runnable on demand via `workflow_dispatch`) sets. Declaration
checks — vacuity, schema conformance — run for both tiers on every push, so a corpus scenario
cannot sit malformed until Monday.

## Scope

Deliberately **not** here, and tracked separately:

- Generated-PDF and visual-regression scoring — that is #443's ratchet, and depending on it before
  it exists would bake unstable numbers into this suite's baselines.
- Agent-scored runs and model planning quality, which consume the scorecard baseline above.
- Delivery change receipts, which belong to the #465 delivery operation (see "Scorecards and
  failure artifacts").

All nine #466 scenario families now run in the fast tier (the whole suite executes in seconds,
so nothing yet warrants the corpus tier, which stands ready for heavier fixtures). Two
deliberate descopes are recorded where they bind:

- `review-annotations` covers comment, threaded reply, footnote, bookmark, and a
  bookmark-anchored internal hyperlink. A field-based `REF` cross-reference is not authorable on
  the tool surface — issue #545 tracks the op; the scenario notes the substitution.
- `layout-preservation` is single-section: no section-break authoring op exists on the tool
  surface, so a multi-section fixture cannot be built through steps. Part identity does the
  layout work — any touch of a header or footer part fails `changedPartsMustBeWithin`.

One engine observation from building the corpus, worth knowing when authoring fixtures: a
comment whose range covers a paragraph carrying live tracked-change markup reports a spurious
`comment` modification in the #457 change set on every open→save cycle (the target paragraph's
content-derived anchor id is not stable across the save normalization). The fixture therefore
anchors its pre-existing comment on a revision-free paragraph.
