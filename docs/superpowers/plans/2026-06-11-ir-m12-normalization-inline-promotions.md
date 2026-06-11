# Document IR — M1.2 Normalization Rules + Typed Inline Promotions

> **For agentic workers:** REQUIRED SUB-SKILL: superpowers:subagent-driven-development. Executed by the controller as sequenced subagent tasks; this file is the persistent task record.

**Goal:** Complete the spec's normalization table (N3, N4, N7, N8, N9, N12, N14, N15) and promote the modeled-but-opaque inline kinds (fields, hyperlinks, note refs, images) to typed IR nodes, with the diagnostic JSON writer extended in lockstep.

**Authority:** `docs/superpowers/specs/2026-06-11-document-ir-spec.md` §5.2 (N-rules), §6.1 (content-hash semantics for the promoted kinds — hyperlink sentinels 0x08/0x09, transparent field results, id-less note refs), §7.3 (inline types — all already exist from M1.1).

**Baseline:** branch `feat/document-ir` @ 9883961 — full suite 1600 green, corpus totality 668/668, 8 golden snapshots.

**Snapshot policy:** every task that changes reader output regenerates the golden snapshots via `DOCXODUS_IR_REGEN_SNAPSHOTS=1` and the diff is REVIEWED (anchors stable, only expected node-shape changes), never blind-committed.

## Task 1: Drops and text mappings — N3, N4, N7, N8, N15(strip)

`IrReader` changes: `w:bookmarkStart`/`w:bookmarkEnd` dropped from the inline stream (N3); `w:lastRenderedPageBreak` dropped (N4); `w:noBreakHyphen` → text U+2011, `w:softHyphen` → text U+00AD (N7, participates in run coalescing); `w:sym` with `@w:char` parseable as hex → text of that codepoint mapped into the PUA-or-direct char with `@w:font` recorded into the run's unmodeled digest container — if unparseable, opaque as today (N8); `w:commentRangeStart`/`w:commentRangeEnd`/`w:commentReference` dropped from the inline stream (N15 strip half; target recording lands with the comments scope in M1.3). Tests per rule incl. hash-stability assertions (bookmarked vs unbookmarked text → equal ContentHash AND FormatFingerprint). Snapshot regen + review.

## Task 2: Field and hyperlink promotion — N9, N14

`w:fldSimple` → `IrFieldRun(@w:instr, children-as-inlines)`. Complex fields: `w:fldChar begin … separate … end` state machine across runs within the paragraph → one `IrFieldRun` (instruction = concatenated `w:instrText`; cached result = inlines between separate and end); malformed/unterminated sequences fall back to opaque inlines for the involved elements (totality). `w:hyperlink` → `IrHyperlink(target, internalTarget: null-for-now, child inlines)` — target resolved from `@r:id` via the main part's hyperlink relationships, or `@w:anchor` for internal links (internal: Target=null, InternalTarget left null in M1.2 — anchor string recorded in… keep `Target = "#" + anchor` convention, document it). ContentHash per spec §6.1 (field transparent; hyperlink 0x08/0x09-framed target + children). New builder constants `SentinelHyperlink = 0x08`, `SentinelHyperlinkTargetEnd = 0x09` in IrHasher. Tests: simple field, complex field, PAGE-field-result-equals-literal-text content hash, hyperlink target participates in ContentHash, hyperlink text not equal to plain text. Snapshot regen + review.

## Task 3: Note refs, images, SDT unwrap — N12 + promotions

`w:footnoteReference`/`w:endnoteReference` → `IrNoteRef` (kind by element, NoteId from `@w:id`); ContentHash = kind sentinel only (id-less, spec §6.1). `w:drawing` containing `a:blip/@r:embed` → `IrInlineImage(partUri, IrHash.Compute(part bytes), extent cx/cy, docPr description)`; any other drawing/pict shape stays opaque. `w:sdt` (block and inline) and `w:smartTag` unwrapped to content (N12): block-level SDT → its `w:sdtContent` blocks walked normally (anchor lands on inner blocks; outer SDT recorded only via provenance); inline SDT → content inlines spliced. Tests: note ref hash id-independence (two docs, different note ids, same text → equal body ContentHash), image promotion fields, image-bytes-hash equality across re-added identical image, SDT-wrapped paragraph equals bare paragraph hashes. Snapshot regen + review.

## Task 4: JSON writer lockstep + completeness guard

Real JSON branches for `IrFieldRun`, `IrHyperlink`, `IrNoteRef`, `IrInlineImage` (shapes: field → kind/instruction/cachedResult-recursive; hyperlink → kind/target/inlines-recursive; noteRef → kind/noteKind/noteId; image → kind/partUri(relative-only)/imageBytesHash/widthEmu/heightEmu/altText). Writer-completeness test: reflection over `IrInline`/`IrBlock` concrete types in the assembly asserting the writer has a non-"unsupported" branch for each (e.g. construct an instance of each and assert serialized kind != "unsupported"). Final snapshot regen + review; CHANGELOG line appended to the existing IR entry.

## Exit criteria

- All N-rules implemented except N15's target-recording half (M1.3) — each with a dedicated test.
- Corpus totality still 668/668; full suite green; snapshots regenerated exactly once per task with reviewed diffs.
- No `"unsupported"` kinds reachable from reader output (completeness test enforces).
