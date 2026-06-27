# DocxDiff Default Migration Plan

> **Status:** Proposed (2026-06-26). This is the execution plan for **decision D4**
> — promoting `DocxDiff` (the IR diff engine) to the blessed default comparison
> API and demoting `WmlComparer` to a documented, still-reachable **legacy**
> engine. It does **not** remove `WmlComparer` and does **not** add `[Obsolete]`
> in this cycle. Companion docs: [`ir_diff_engine.md`](ir_diff_engine.md) (the
> engine), [`comparison_engine.md`](comparison_engine.md) and
> [`wml_comparer_gaps.md`](wml_comparer_gaps.md) (the incumbent),
> [`docxdiff_libreoffice_findings.md`](docxdiff_libreoffice_findings.md) (oracle).

## 1. Goal & non-goals

**Goal.** Make `DocxDiff` the default engine that every Docxodus surface
(.NET core, WASM, npm, python, CLI) reaches for when a caller does not ask for a
specific engine, and re-label `WmlComparer` as legacy — kept, supported, and
selectable, but no longer blessed.

**Non-goals (this cycle).**

- Removing `WmlComparer` or any of its public surface.
- Adding `[Obsolete]` attributes (deferred to a future major — see §5, M-F).
- Changing the IR diff engine's public API shape. The flip is a *default and
  documentation* change plus the blocking parity work in M-A; it is not an
  engine rewrite.

## 2. Current state — why this is close, and what blocks it

### 2.1 Measured parity (already green, ratcheted in CI)

The pairwise path is at or beyond parity, enforced by ratchet tests in
`Docxodus.Tests/Ir/Diff/`:

| Scoreboard | Measures | Floor | Current |
|---|---|---|---|
| `IrParityScoreboardTests` | `GetRevisions` count/type/text, exact to `WmlComparer` | `GenuinePassFloor = 179` | 179/179 PASS, `Deviation == 0` |
| `IrMarkupParityScoreboardTests` | produced-markup round-trip (accept≡right, reject≡left, schema-valid) | `MarkupParityFloor = 39` | 39/39; 1 allowlist fixture where the `WmlComparer` oracle itself throws |
| `ConsolidateParityScoreboardTests` | N-way consolidate soundness over the 84 legacy corpus cases | `ReproduceFloor = 84` | 84/84, 0 fail |
| `IrVsWmlComparerTests` | broad 92-pair × 2-direction sweep | `CompatMatchFloor = 150`, `DivergentCeiling = 66`, genuine-loss causes capped at 0 | held; divergent tail fully bucketed into benign / "IR is more correct" causes |

`DocxDiff` is also **strictly ahead** on: comments surviving edits (the
`WmlComparer` oracle drops comments on any edit), anchor-addressed revisions
(`kind:scope:unid`), diff-as-data JSON, and determinism by construction.

### 2.2 What blocks the flip today

1. **Four multi-reviewer `Consolidate` edge-gaps** (see M-A). These are the only
   places where `DocxDiff` does *less* than the incumbent. They fail loud (no
   silent data loss), but they must be closed before the flip — this is a hard
   prerequisite, not an accepted-scope reduction.
2. **Asymmetric client wiring.** The shipped clients are split-brained:

   | Surface | Engine wired today | File |
   |---|---|---|
   | WASM `DocumentComparer` | `WmlComparer` only | `wasm/DocxodusWasm/DocumentComparer.cs` |
   | npm `compareDocuments*` | `WmlComparer` only | `npm/src/index.ts` |
   | `tools/redline` CLI | `WmlComparer` only | `tools/redline/Program.cs` |
   | WASM `DocxDiffBridge` / npm `docxDiff*` | `DocxDiff` only | `wasm/DocxodusWasm/DocxDiffBridge.cs`, `npm/src/index.ts` |
   | python `docx-scalpel` | `DocxDiff` only | `python/src/docx_scalpel/session.py` |
   | `tools/diffharness` | `DocxDiff` only | `tools/diffharness/DiffRunner.cs` |

   No surface lets a caller pick an engine, and no two surfaces agree on a
   default. M-B fixes this.
3. **Sign-off gates still open.** `G2` (native-markup-renderer go/no-go) and
   `D4` (default flip) are both unratified, gated on a Microsoft Word
   manual-verification pass (M-C) and a burn-in window (M-D).
4. **Two honest default deviations** from `WmlComparerSettings` that a drop-in
   default must consciously own (see §8): `Deterministic` dates default `true`,
   and `FormatComparison` defaults to `ModeledOnly` (vs. `WmlComparer`'s `Full`).

## 3. Guiding principles

- **Strangler, not rewrite.** `WmlComparer` keeps working untouched and stays
  reachable for the entire migration and after it. The flip changes which engine
  is reached *by default*, nothing else.
- **Oracle-driven.** Every M-A gap ships its correctness oracle (a scoreboard
  ratchet + round-trip invariant + headless-LibreOffice load backstop) before it
  is called done. No gap closes on assertion alone.
- **One source of truth for the default.** Per the "same default everywhere"
  decision, the blessed engine is chosen in exactly one place in the core
  library; every client's default *mirrors* that value rather than pinning its
  own. Flipping the default is then a one-line change in one file.

## 4. The engine-selection model ("same default" design)

The cutover introduces an explicit, selectable engine with a single shared
default. **`WmlComparer` stays the default for now** — M-B adds the selector but
does **not** flip the default; the default changes to `DocxDiff` only at M-E,
once the M-A/M-C/M-D gates clear. The model:

- A core-library notion of "the default comparison engine" lives in **one**
  location — e.g. a `ComparisonEngine { DocxDiff, WmlComparer }` enum plus a
  single `ComparisonDefaults.Engine` constant. Before the flip it resolves to
  `WmlComparer`; the flip changes it to `DocxDiff`. Nothing else hard-codes a
  default.
- Each client surface gains an **optional** engine selector whose default is
  *unset*, meaning "use the core default":
  - WASM `DocumentComparer.*` / `DocxDiffBridge.*` → an `engine` argument.
  - npm `compareDocuments*(…, { engine })`.
  - `tools/redline` → `--engine docxdiff|wmlcomparer`.
  - python `docx-scalpel` already resolves to the target default; expose the
    legacy selector only if a consumer needs it (low priority — see M-B).
- Because clients inherit rather than pin the default, the parity matrix and the
  rollback story are both governed by that single constant.

This is the rollback mechanism too: a regression after the flip is mitigated by
selecting `wmlcomparer` per-call, with no redeploy of engine code.

## 5. Milestones

Sequenced. M-A is the blocking engineering work; M-B/M-C run in parallel with it;
M-D depends on all three; M-E is the flip; M-F is deferred.

### M-A — Close the four `Consolidate` edge-gaps (BLOCKING)

These are the documented v1 limitations in `ir_diff_engine.md` and
`Docxodus/Ir/Diff/IrCompositeMerger.cs`. Each closes with a scoreboard ratchet +
reject≡base / accept≡policy-composite round-trip + LibreOffice load backstop.

| # | Gap | Today | Target |
|---|---|---|---|
| A1 | **Note-scope N-way merge** | N≥2 consolidate throws `NotSupportedException` when any reviewer edits footnote/endnote content (`IrCompositeScript.NoteOps` always null). | Compose note-scope ops across reviewers exactly as body ops; lift the tripwire to a real merge path. |
| A2 | **Structural table changes across reviewers** | `MovedRow`, column-count change, and cell-shell-only (`tcPr`/`pPr`) changes fall back to whole-table block conflict. | Per-cell compose across row moves and column add/remove; surface column add/remove as a tracked table change rather than an invisible/blocked one. |
| A3 | **Cross-reviewer split/merge** | `SplitBlock`/`MergeBlock` are lowered to del/ins inside consolidate (markup not preserved across reviewers). | Preserve native split/merge markup across reviewers, with per-reviewer attribution. |
| A4 | **Cross-kind note nesting** | An endnote reference inside a footnote body (or vice-versa) is not renumbered. | Renumber cross-kind nested references during the note renumber walk. |

**Exit criteria:** all four behaviors land with new multi-reviewer fixtures;
`ConsolidateParityScoreboardTests` floor raised to cover them; the
`NotSupportedException` tripwire in `IrCompositeMerger` removed (or downgraded to
a genuinely-unsupported residual that is documented and out of the four); full
`.NET` suite + Release (warnings-as-errors) green.

### M-B — Unify the engine selector and cut clients over

1. Introduce the single-source-of-truth default (§4) in core, **seeded to
   `WmlComparer`** — M-B does not change the default, it only makes the engine
   selectable; the flip to `DocxDiff` is M-E. If a shared comparison facade is
   needed so both engines sit behind one entrypoint (there is no `WmlComparerOps`
   facade today — WASM calls `WmlComparer` directly), add a thin one next to
   `DocxDiffOps` and route both bridges through it.
2. Thread the `engine` selector through every client surface in §2.2, defaulting
   to the core default. Follow the four-layer ripple (§7).
3. Add a `tools/redline --engine` flag and a `both`/diff mode that runs both
   engines and reports divergence (feeds M-D burn-in).

**Exit criteria:** every client can select an engine; with no selection, every
client resolves to the same core default; npm `tsc` + Playwright green; python
e2e green; `tools/redline` smoke green on both engines.

### M-C — Author and run the Word manual-verification checklist

The checklist artifact does **not** exist yet. Create
`docs/architecture/word_verification_checklist.md` (committed) covering, per a
representative corpus (incl. `TestFiles/WC/*`, `TestFiles/DD/*`, and a few real
contracts):

- Insertions/deletions render and **accept → right**, **reject → left** in Word.
- Moves render as `w:moveFrom`/`w:moveTo` (not del/ins) when `DetectMoves`.
- `w:rPrChange` format changes render and round-trip.
- Comments survive an edit and remain threaded.
- Footnote/endnote references renumber and resolve.
- Nested + structural tables.
- A sign-off table: Word build (Windows 365, Windows 2019, Mac), reviewer,
  date, pass/defects.

**Exit criteria:** checklist authored; one full human pass executed with zero
Severity-1 (data-loss / fails-to-open / accept≠right) defects outstanding.

### M-D — Burn-in

- Promote the existing `IrVsWmlComparerTests` differential to a standing
  per-release signal and add the `tools/redline --engine both` divergence report
  to the burn-in corpus run.
- **Pass criteria:** at least one full release cycle (≥ 2 weeks) of dual-run with
  (a) zero new divergence-regressions beyond the ratcheted ceilings, (b) zero
  open Severity-1 Word-verification defects from M-C, (c) `InspectCompatibility`
  surfacing no un-triaged `Untested` feature on the burn-in corpus.

### M-E — The flip

1. Change the single shared default (§4) from `WmlComparer` to `DocxDiff`.
2. Re-label legacy **in docs and `///` remarks only** (no `[Obsolete]`):
   - `Docxodus/DocxDiff.cs` remarks (drop "prefer `WmlComparer` for production").
   - `WmlComparer` class `///` → "legacy; prefer `DocxDiff`".
   - `ir_diff_engine.md` status line + the relationship table row.
   - `CLAUDE.md` — replace "`WmlComparer` remains the default/blessed" language.
3. `CHANGELOG.md` `[Unreleased]` → a prominent **Changed** entry: "default
   comparison engine is now `DocxDiff`; `WmlComparer` remains available via the
   `engine` selector." (Semver call in §8.)
4. Mark **G2** and **D4** ratified — the authoritative status now lives in the
   tracked files (`ir_diff_engine.md` status line + `DocxDiff.cs` remarks);
   update the internal decision log too if it is still kept.

**Exit criteria:** default resolves to `DocxDiff` on every surface; the legacy
selector still produces `WmlComparer` output; docs/CHANGELOG updated; gates
ratified.

### M-F — `[Obsolete]` (deferred)

At a future major release, add `[Obsolete]` (warning, not error) to
`WmlComparer`'s public entry points. Out of scope for this cycle by decision.

## 6. Gate exit criteria

| Gate | Meaning | Exit criteria |
|---|---|---|
| **G2** | Native-markup-renderer viability | `IrMarkupParityScoreboardTests` green at floor; LibreOffice backstop green; M-C markup section signed with no Sev-1. |
| **D4** | `DocxDiff` becomes default | M-A complete (4 gaps closed + scoreboards ratcheted); M-B complete (selector everywhere, one default); M-C signed; M-D burn-in passed. |

## 7. Four-layer ripple checklist (per CLAUDE.md)

For the M-B selector and the M-E default change:

- [ ] **Core:** `ComparisonEngine` enum + single default constant; optional
      shared comparison facade routing both engines.
- [ ] **Tests:** selector-resolution tests; default-resolves-to-DocxDiff test;
      M-A scoreboard ratchets.
- [ ] **WASM bridge:** `engine` arg on `DocumentComparer` / `DocxDiffBridge`.
- [ ] **npm/TS:** `engine` option on `compareDocuments*`; `types.ts` enum.
- [ ] **python-host + docx-scalpel:** selector parity (legacy optional).
- [ ] **CLI:** `tools/redline --engine` + `both` divergence mode.
- [ ] **Docs:** this plan, `ir_diff_engine.md`, `comparison_engine.md`,
      `CLAUDE.md`, CHANGELOG, the new Word checklist.

## 8. Risks, decisions, and rollback

| Item | Risk / decision | Disposition |
|---|---|---|
| `FormatComparison` default | `ModeledOnly` can under-report visible-but-unmodeled format changes (e.g. `w:shd` run shading) vs. `WmlComparer`'s `Full`. | **Decision point at M-E:** keep `ModeledOnly` (recommended — less noise; document the deviation) or set `Full` for byte-drop-in. |
| `Deterministic` dates default | `true` (pinned epoch) vs. `WmlComparer`'s wall-clock. | Keep — it is an improvement; document in the CHANGELOG. |
| Semver of the flip | Changing the default engine changes output for callers who don't pass a selector. | **Decision point:** treat as **major** (default behavioral change) — recommended — or minor since legacy stays opt-in. |
| Real-world feature coverage | A production doc hits an `Untested`/`Partial` feature. | `DocxDiff.InspectCompatibility` pre-flight + M-D burn-in; legacy selector as escape hatch. |
| **Rollback** | A post-flip regression in the field. | Select `engine = wmlcomparer` per call — no engine redeploy. The selector *is* the rollback. |

## 9. Definition of done — "WmlComparer is officially legacy"

- [ ] M-A: all four `Consolidate` gaps closed; scoreboards ratcheted; tripwire removed.
- [ ] M-B: engine selector on every surface; one shared default; all client suites green.
- [ ] M-C: Word checklist authored and signed, zero open Sev-1.
- [ ] M-D: burn-in window passed against its criteria.
- [ ] M-E: default flipped; legacy reachable; docs/CHANGELOG updated; G2 + D4 ratified.
- [ ] Deferred (M-F): `[Obsolete]` scheduled for the next major.
