# Diff Engine — M2.4b Deviation Burndown

> **For agentic workers:** REQUIRED SUB-SKILL: superpowers:subagent-driven-development. Executed by the controller as sequenced subagent tasks.

**Goal:** Methodically close the residual parity gaps: the **18 documented deviations** on the GetRevisions scoreboard and the **8 allowlisted fixtures (6 root causes)** on the markup round-trip — they share root causes, so workstreams are organized by cause, not by row.

**USER METHOD RULE (binding):** per gap: (1) reproduce and diagnose from first principles; (2) the INITIAL PRESUMPTION is that WmlComparer's expected behavior is CORRECT — the IR engine should be fixed to match; (3) only if investigation ESTABLISHES WmlComparer is wrong (with concrete evidence: spec citation, Word behavior, or demonstrable self-inconsistency) does the row stay a documented deviation — with the strengthened evidence attached. "IR is finer/cleaner" is NOT sufficient grounds; matching behavior wins unless the oracle is proven faulty.

**Standing constraints:** Fine-mode edit-script grain remains available (render-time policies are fine; engine fixes that IMPROVE correctness are fine; coarsening the engine's data model is not). Consolidate-compatible architecture. Ratchets only rise: GetRevisions floor 179 (composition shifts deviation→PASS), markup floor 39, allowlist only shrinks.

**Baseline:** `feat/diff-m24` @ 710b26f — 218/218 (161 PASS + 18 DEV on GetRevisions; 39/39 markup; 8-fixture allowlist). NO merges to main during this milestone (user instruction).

## Workstream A: Relationship-id-stable opaque hashing (closes WC-1940 + markup WC014/WC052 SmartArt ×3, WC022 image/math)

Root cause: `IrHasher.Canonicalize` hashes raw `r:id`/`r:embed` attribute VALUES, so a renumbered-but-identical diagram/image reads as changed. Fix at the reader/hasher level: during canonicalization (and the opaque-inline hash path), resolve relationship-id attributes to a STABLE token — target part content hash for internal parts (cache per part), target URI for external rels, sentinel for dangling. Snapshot churn expected (reviewed; only opaque hashes change). Verify: WC-1940 scoreboard row → genuine PASS; the 4 markup allowlist fixtures → round-trip clean, allowlist shrinks; markdown-projection equivalence suite must stay green (the emitter never surfaces opaque hashes in output — verify); full corpus totality.

## Workstream B: Low-similarity Modified rendering (the "coincidental Equal island" family — WC-1170, 1190, 1210, 1420, 1430, 1440, 1450, 1830, 1840, 1950, and WC-1770's textbox variant)

Diagnosis to confirm per-case: the 1×1-gap fallback (and similar) pairs effectively-rewritten paragraphs as Modified; Myers then credits coincidental shared words ("Video") as Equal islands → more revisions than WmlComparer's clean whole-region del+ins. PRESUMPTION CHECK: for a true rewrite, WmlComparer's 2-revision report is the better account — expect most/all of this family to be IR-side fixes. Fix candidate (render-time, compatible mode at minimum — investigate whether Fine should share it): when a ModifyBlock's Equal+FormatChanged token coverage is below a threshold (derive empirically; e.g. <25% of either side), render whole-block del+ins. Per-case verdicts table required (each of the 11 rows: diagnosis, which engine right, fix applied or evidence). Target: family → genuine PASS.

## Workstream C: Structural pairing gaps (WC-1750/1760 endnote tables; WC-1710/1720 note-ref attribution; WC034 markup pair)

- WC-1750/1760: the aligner never pairs the two endnote tables as Modified (IrBlockSimilarity scores non-paragraphs 0) → whole-table del+ins (3) vs WmlComparer's row-level (6). PRESUMPTION: WmlComparer right. Engine improvement (allowed — adds capability): table-aware similarity (score over concatenated cell token multisets) so in-gap pairing can produce Modified table pairs feeding IrTableDiffer. Verify no regression on table corpus rows.
- WC-1710/1720 + WC034: disentangle the TWO effects previously bundled: (a) WmlComparer reporting unchanged body word as del+ins under note-renumber — INVESTIGATE: is that a real oracle fault (evidence!) or correct-by-its-semantics? (b) IR affix-trim coalescing. Re-derive expected counts after Workstream B lands (interactions likely). Verdicts + evidence per the method rule.

## Workstream D: Duplicate-textbox dedup (WC-1900/1920) + markup leftovers (WC019 hyperlink rId remap, WC-BodyBookmarks markers) + close

- WC-1900 (separate-cell Choice/Fallback duplicates) + WC-1920 (nested): the reader emits BOTH mc:AlternateContent copies (matches the markdown oracle — constraint: M1.4 equivalence must stay green, so dedup must NOT change reader output for the projection path). Investigate dedup at the DIFF layer (extend the pair-walk dedup to non-adjacent/nested cases) per the presumption that WmlComparer's single-count is right.
- WC019: true rId remap in IrMarkupRenderer (rewrite cloned @r:id + recreate rel under a fresh id on collision).
- WC-BodyBookmarks: bookmark/perm marker revisions — investigate what WmlComparer produces and whether matching requires reader bookmark modeling (N3 currently drops them); scope honestly — if it needs bookmark-range modeling, propose as M2.5 item with the design sketch rather than rushing it.
- Close: final scoreboard composition + allowlist end-state; updated deviation catalog (every surviving deviation carries established-oracle-fault evidence); `## M2.4b Outcome`; program plan + CHANGELOG; full verification.

## Exit criteria

- Every one of the 18+8 gaps has a per-row verdict: FIXED (engine/render — genuine PASS) or ORACLE-FAULT-ESTABLISHED (deviation retained with concrete evidence).
- Ratchets risen to match; allowlist minimized; full suite + corpus + fuzz + projection equivalence green; no main merges.
