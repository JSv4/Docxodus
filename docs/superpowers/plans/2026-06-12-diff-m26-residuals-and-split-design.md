# Diff Engine — M2.6 Residual Closure + 1:N Design Hardening

> **For agentic workers:** REQUIRED SUB-SKILL: superpowers:subagent-driven-development. Controller-sequenced tasks.

**User direction:** close the non-1:N residuals FIRST; then harden the 1:N split/merge sketch into a fully resolved, reviewed design — **implementation of 1:N stays deferred**. Same method rules: oracle presumed correct, evidence to deviate, Fine grain preserved, LOCAL ONLY, no main merges, granular commits, per-task user updates.

**Baseline:** `feat/diff-m24` @ (T5 docs commit). GetRevisions 177+2/179; markup 39/39; allowlist 4 (WC034 ×2, WC022, WC-BodyBookmarks); suite 1954/0.

## Task 1: Note-id renumber output pass (closes WC034 ×2)

The markup renderer's produced document keeps LEFT-package note ids; WC034's expected round-trip needs the oracle-equivalent of `ChangeFootnoteEndnoteReferencesToUniqueRange`: renumber footnote/endnote ids in the OUTPUT package to body-reference document order (definitions reordered/renumbered to match, refs rewritten, separator/continuation boilerplate ids preserved per OOXML rules — study the oracle's pass + what RevisionProcessor expects). Verify: WC034 ×2 leave the allowlist (4→2); round-trip + validation + note-scope invariants corpus-wide; old-engine suite untouched.

## Task 2: WC022 ordering + Deleted-anchor nuance + WC-BodyBookmarks verdict

- WC022 residual: "adjacent-empty-paragraph alignment ordering sensitivity" — reproduce, diagnose precisely, fix if bounded (deterministic ordering preference) else evidence + retain (it's an aligner-order quirk, possibly M2.6-design-adjacent — say so if it reduces to 1:N).
- Deleted-revision-carries-both-anchors (python E2E observation): establish the intended anchor-presence-by-type contract (Deleted = left-only? Moved = one per side?), fix the renderer or the doc — small but the public surface just shipped, so get it right + test.
- WC-BodyBookmarks: final verdict per method rule. The ORACLE THROWS on this fixture ("Internal error", endnote→footnote whole-store conversion) — so there is no oracle behavior to match. Decide with evidence: either our engine handles it (round-trip clean = we EXCEED the oracle — acceptable, document) or retain with the oracle-throws evidence as the ceiling. Do not sink unbounded time; the verdict + documentation is the deliverable.

## Task 3: 1:N split/merge design hardening (NO implementation)

Take `docs/superpowers/specs/2026-06-12-subparagraph-split-merge-design.md` to a resolved design spec: exact edit-script op shapes (`IrSplitBlockOp`/`IrMergeBlockOp` or n-ary `IrEditOp` generalization — decide with rationale), detection algorithm (containment thresholds, determinism, cost bounds, interaction with similarity pairing + move detection), apply-verifier contract, markup-renderer emission (what does the ORACLE's output look like for the WC-1450/1830 splits — derive the target shapes from its actual produced markup), revisions-surface + JSON + python/npm wire impact, consolidate-compatibility check, risk register, test plan, and an explicit implementation-effort estimate. Then an adversarial DESIGN REVIEW (separate reviewer agent) challenging: op-model soundness, 1:N vs N:1 vs N:M scope creep, the apply-verifier semantics, and regression surface. Output: revised spec marked DESIGN-RESOLVED + review findings appended. NO code.

## Exit

Allowlist ≤2 with verdicts everywhere; anchor contract fixed+tested; 1:N spec DESIGN-RESOLVED with adversarial review; suite green; everything local on the feature branch.
