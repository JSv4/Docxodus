# Diff Engine — M2.4 Native Markup + Parity Closure (GO/NO-GO GATE)

> **For agentic workers:** REQUIRED SUB-SKILL: superpowers:subagent-driven-development. Executed by the controller as sequenced subagent tasks.

**Goal:** Drive the parity scoreboard from 129/218 to **218/218** — the M2.4 gate per the user directive. Four workstreams: scope-complete diffing, render-time granularity compatibility, the native OOXML revision renderer (the program's historical risk concentration), and the gate report (decision G2).

**Baseline:** `feat/diff-m24` @ d6a20e7. Scoreboard 129/179 runnable (ratchet floor 129); 39 markup-blocked.

**USER ADJUDICATIONS (binding):**
- Consolidate OUT of scope; bar = 218/218. Do NOT architecturally preclude it: revision/markup surfaces stay multi-author-capable (author is per-revision data, not a singleton assumption); no type bakes in "exactly two documents, one comparison" where WmlComparer's consolidate would need N.
- Granularity compatibility is RENDER-TIME POLICY ONLY: `IrDiffSettings.RevisionGranularity = WmlComparerCompatible | Fine`. The edit script's fine grain is untouchable. The adapter/scoreboard uses WmlComparerCompatible; Fine remains the engine default... decide: default = Fine (the engine's native truth; compatible mode is an explicit opt-in for parity consumers) — document.

**Ratchet discipline:** every task that closes scoreboard cases RAISES the `ParityFloor` constant in IrParityScoreboardTests in the same commit. The floor never goes down; 218 at gate.

## Task 1: Scope-complete diffing (footnotes/endnotes/textboxes)

Aligner gains scope coverage: align footnote/endnote stores (note-id-matched scopes → block alignment within each matched note; unmatched note = whole-note insert/delete) and textbox inner blocks (IrTextbox placeholder tokens already flag the change at paragraph grain; the edit script needs the INNER block diff: when a Modified paragraph pair's textbox placeholders differ, recurse alignment over the paired textboxes' Blocks — design: nested ops attached to the ModifyBlock, mirroring the TableDiff pattern). Revisions render from all scopes (note revisions carry the note scope context — fn/en anchors already distinguish them). Headers/footers: WmlComparer compares... CHECK what it does (its GetRevisions covers footnotes/endnotes via ProcessFootnoteEndnote; headers likely not — mirror its scope coverage for parity, and include headers only if any test demands it). Apply-verifier + invariants extended to the new scopes. Expected closure: the 21 scope-gap scoreboard failures + the 34-pair ScopeGapNewEmpty differential bucket. Raise floor accordingly.

## Task 2: Render-time granularity + DetectMoves switch

- `RevisionGranularity { Fine (default — engine truth), WmlComparerCompatible }` consumed by IrRevisionRenderer ONLY: under compatible mode, coalesce adjacent same-kind token-op revisions within a block (WmlComparer emits per contiguous changed region), suppress context re-attribution (tune: when an in-gap similarity pairing produced wider-than-minimal spans, compatible mode trims common prefix/suffix tokens between the paired del/ins span texts — render-time trimming, document precisely), and match WmlComparer's revision-text composition (study failing WC003 cases' expected counts to derive the exact rules — iterate against the scoreboard).
- DetectMoves=false: renderer (not aligner) demotes Moved/MovedModified to Inserted+Deleted pairs; adapter maps WmlComparerSettings.DetectMoves accordingly. Engine alignment unchanged (capability preserved per directive).
- The 3 cases where WmlComparer under-reports (WC-1960/70/80 French apostrophe family): tune compatible mode to reproduce WmlComparer's expected counts ONLY if achievable without corrupting Fine mode; if reproducing requires emulating the old engine's bug, document the bug, mark the ported expectation with the deviation, and count them as scoreboard passes via documented-expected-difference rows — controller adjudicates the final call on these 3 at review.
- Expected closure: the 27 granularity + 2 move-switch failures → floor reaches 179/179 runnable.

## Task 3: Native OOXML revision renderer — core (w:ins/w:del)

`IrMarkupRenderer.Render(IrEditScript, left, right, IrDiffSettings) → WmlDocument`: produce the compared document — right document's content with tracked-change markup expressing the script (the WmlComparer output contract: accept-all ⇒ right, reject-all ⇒ left). Build on the LEFT document's package (styles/numbering/parts continuity — study how WmlComparer assembles its output and what the markup-blocked tests assert: read them ALL before designing). Core scope this task: body paragraphs — EqualBlock passthrough; InsertBlock → content wrapped in w:ins; DeleteBlock → left content as w:del with w:delText; ModifyBlock → token spans to runs split at span boundaries, w:ins/w:del per span (run properties preserved from the source tokens' formats; rebuild runs from IR tokens); revision ids unique + deterministic (single ascending counter per render — NO static state); author/date from settings. Validation: OpenXmlValidator clean (or document accepted warnings); THE INVARIANT (fuzz + corpus): RevisionProcessor.AcceptRevisions(output) content-equals right (IR ContentHash comparison per block) AND RejectRevisions(output) content-equals left. Tables/format/moves/notes = Task 4 (emit conservative whole-block ins/del for unsupported constructs so the invariant still holds — document interim coarseness). Run the markup-blocked scoreboard rows that only need ins/del; report which pass.

## Task 4: Native renderer completion + gate (G2)

w:moveFrom/w:moveTo with move ranges (+ DetectMoves/SimplifyMoveMarkup settings parity); w:rPrChange for FormatChanged spans; table row/cell revisions (trPr/ins, del, cell markers per the markup-blocked tests' expectations); footnote/endnote scope markup; whatever LegalNumbering/BodyLevelElements/BodyLevelBookmark/ParallelRace assert (ParallelRace: the IR engine must be thread-safe by construction — prove with the ported test). Drive the scoreboard to 218/218; floor = 218. `## M2.4 Outcome` = the GATE REPORT: scoreboard final, accept/reject invariant fuzz results, validation status, Word-manual-check recommendation list for the user, G2 GO/NO-GO recommendation + D4 (default engine) considerations for M2.5. Program plan + CHANGELOG.

## Out of scope

Consolidate (per adjudication; architecture kept compatible). Public API surface (M2.5). SimplifyMoveMarkup beyond what tests assert.
