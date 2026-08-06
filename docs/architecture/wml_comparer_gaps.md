# WmlComparer.cs - Gaps and Deficiencies

This document catalogs known gaps, limitations, and areas for improvement in the WmlComparer (document comparison engine).

> **Update (v6.x / M2.5):** Several gaps below were closed in the v6.x line and are no longer accurate — the corrections are inline. Most are also addressed structurally by the new **IR diff engine** (`DocxDiff`), which is anchor-addressed and emits the edit script as data; see [`ir_diff_engine.md`](./ir_diff_engine.md). `WmlComparer` remains the default/blessed comparison API.

## 1. Limited Revision Types Exposed

**Location:** `WmlComparer.cs` lines 3317-3321

The `WmlComparerRevisionType` enum only exposes two revision types:

```csharp
public enum WmlComparerRevisionType
{
    Inserted,
    Deleted,
}
```

However, the **HTML renderer** (`WmlToHtmlConverter`) supports rendering additional tracked change types from Word documents:
- `w:ins` - Insertions (rendered as `<ins>`)
- `w:del` - Deletions (rendered as `<del>`)
- `w:moveFrom` / `w:moveTo` - Move operations (when `RenderMoveOperations` is enabled)
- `w:rPrChange` / `w:pPrChange` - Format changes (described via `DescribeFormatChange`)

### Impact

When comparing two documents, the `GetRevisions()` API cannot distinguish between:
- Content that was moved vs. deleted and re-inserted elsewhere
- Pure text changes vs. formatting-only changes

### Internal State Not Exposed

The internal `CorrelationStatus` enum has more granular states that are not exposed to consumers:
- `Nil`, `Normal`, `Unknown`, `Inserted`, `Deleted`, `Equal`, `Group`

### Recommendation

Extend `WmlComparerRevisionType` to include:
- `Moved` - For content detected as moved (currently shows as deletion + insertion pair)
- `FormatChange` - For formatting-only modifications

This would bring the comparison API in line with what the HTML renderer already supports for documents with existing tracked changes.

## 2. Move Detection ✅ IMPLEMENTED

**Status:** Implemented (Issue #20)

Move detection has been implemented in `GetRevisions()` using post-processing with Jaccard similarity:

- **`WmlComparerRevisionType.Moved`** - New enum value for moved content
- **`WmlComparerRevision.MoveGroupId`** - Links source and destination revisions
- **`WmlComparerRevision.IsMoveSource`** - true=moved FROM here, false=moved TO here
- **Settings**:
  - `DetectMoves` (default: true)
  - `MoveSimilarityThreshold` (default: 0.8 = 80% word overlap)
  - `MoveMinimumWordCount` (default: 3 words minimum)

The implementation uses word-level Jaccard similarity to match deletions with insertions. When similarity exceeds the threshold and the text meets minimum word count, the pair is marked as a move.

**Note (corrected, v6.x):** Earlier revisions of this doc claimed native `w:moveFrom`/`w:moveTo` markup is NOT generated. **That is stale.** `WmlComparer.Compare` produces native Word move markup (`w:moveFromRangeStart`/`w:moveFromRangeEnd`/`w:moveToRangeStart`/`w:moveToRangeEnd` + `w:moveFrom`/`w:moveTo` keyed by `w:name`) when `DetectMoves` is enabled, and `GetRevisions()` recognizes it (see CLAUDE.md "WmlComparer.cs"). The new `DocxDiff` IR engine likewise emits native move markup ([`ir_diff_engine.md`](./ir_diff_engine.md)).

## 3. Format Change Detection ✅ IMPLEMENTED

**Status (corrected, v6.x):** Implemented. Earlier revisions called this a gap; **that is stale.** `WmlComparer` exposes a `FormatChanged` revision type with `DetectFormatChanges` (default true): when text is identical but modeled run formatting differs, it produces native `w:rPrChange` markup and `GetRevisions()` returns a `FormatChanged` revision whose `FormatChange` carries old/new properties and changed names (see CLAUDE.md "WmlComparer.cs" and `docs/architecture/format_change_detection.md`). The new `DocxDiff` IR engine surfaces the same via `DocxDiffRevisionType.FormatChanged` + `DocxDiffFormatChange`, with a `FormatComparison` (ModeledOnly | Full) policy ([`ir_diff_engine.md`](./ir_diff_engine.md)).

Word documents can track formatting changes via:

| Markup | Scope | WmlComparer |
|--------|-------|-------------|
| `w:rPrChange` | run properties | ✅ `DetectFormatChanges` (default **true**) |
| `w:pPrChange` | paragraph properties — alignment, indent, spacing, style, numbering | ✅ `DetectParagraphFormatChanges` (default **false**) |
| `w:sectPrChange` | section properties | ✗ |
| `w:tblPrChange` | table properties | ✗ |

### 3a. Paragraph properties ✅ IMPLEMENTED

`WmlComparerSettings.DetectParagraphFormatChanges` extends the run-level pass to the paragraph mark.
Off by default: enabling it changes the output of comparisons that produce no revision today.

Why it is worth having on: without it a paragraph-format-only change is *silently lossy*. The
paragraph mark hashes as `"pPr"` regardless of its contents (the hash is element name + text value,
and every `w:pPr` child is attribute-only), so such a paragraph correlates as Equal and the result
is rebuilt from the RIGHT side. The revised formatting is therefore already applied to the output,
the original's is discarded, and nothing is reported — so reject cannot put it back.

Mechanics, mirroring the run-level pass exactly:

- `DetectFormatChangesInAtomList` branches on a `w:pPr` atom and compares the two sides through
  `ReduceToParagraphPropertiesChange`, which also produces the archived copy — what is compared is
  precisely what is stored.
- The reduction drops what a `w:pPrChange` may not carry: its inner `w:pPr` is a **CT_PPrBase**, so
  `w:rPr` (the mark's own run properties, which Word tracks separately) and `w:sectPr` are excluded,
  along with any pre-existing `w:pPrChange`. Revision-save ids and `pt14:` bookkeeping are stripped
  at every level, so they can neither cause a false positive nor leak into the archive.
- The reduction sorts properties at every level, so the same properties written in a different
  source order — including a `w:numPr` holding `w:numId` before `w:ilvl` — do not read as a change;
  `XNode.DeepEquals` is order-sensitive. That sort is not the schema sequence, but the writer's
  element ordering restores it in the emitted archive (pinned by `PF004b`).
- Accept/reject needed no new code — `RevisionProcessor` already handles `w:pPrChange`, including
  carrying an inline `w:sectPr` across a reject.

**Lists come free.** `w:numPr` is an ordinary `w:pPr` child, so list membership, level and format
changes are tracked by the same path. Renumbering caused by editing `numbering.xml` — changing what
a `numId` *means* rather than which one a paragraph points at — is not a `w:pPr` change and is not
tracked. Word does not track that either.

Known limits:

- The paragraph mark's own run properties (`w:pPr/w:rPr`) are not tracked, and neither is a section
  change.
- **Body scope only.** No formatting change inside a footnote or endnote is detected — and this is
  pre-existing, not a property of the new option: the run-level pass misses note-scope `w:rPr`
  changes in exactly the same way, while a *text* change in the same note is detected normally, so
  notes are compared and it is the format passes that do not reach them. `GetFormatChangeRevisions`
  scans the note parts, so the reporting side is ready if detection is ever extended. Pinned by
  `PF022`, which fails if that changes.
- WmlComparer discards a mid-document inline `w:sectPr` altogether — again pre-existing and
  unrelated to this option, and the reason the CT_PPrBase exclusion is covered by a direct unit test
  rather than through `Compare`.

Coverage: `WmlComparerParagraphFormatTests`. CLI: `redline --detect-paragraph-format-changes`.

## 4. Revision Metadata Limitations

**Location:** `WmlComparerRevision` class

The current revision class exposes:
```csharp
public class WmlComparerRevision
{
    public WmlComparerRevisionType RevisionType;
    public string Text;
    public string Author;
    public string Date;
    public XElement ContentXElement;
    public XElement RevisionXElement;
    public Uri PartUri;
    public string PartContentType;
}
```

### Missing Information

- **Move pair linking** - If moves were detected, there's no way to link the "from" and "to" revisions
- **Paragraph context** - The surrounding paragraph or heading for context
- **Position information** - Character offset or paragraph number in the document

## 5. npm/TypeScript API Reflects .NET Limitations

The TypeScript `RevisionType` enum mirrors the .NET limitation:

```typescript
export enum RevisionType {
  Inserted = "Inserted",
  Deleted = "Deleted",
}
```

When the .NET comparison engine is enhanced, the TypeScript types should be updated accordingly:
- `npm/src/types.ts` - Add new enum values
- `npm/src/index.ts` - Update exports
- `wasm/DocxodusWasm/DocumentComparer.cs` - Update WASM bridge

---

## Summary of Priority Improvements

### Completed ✅

1. ~~**Add move detection**~~ - ✅ Implemented: `Moved` revision type with `MoveGroupId` and `IsMoveSource` properties
2. ~~**Link related revisions**~~ - ✅ Move pairs are now linked via `MoveGroupId`
3. ~~**Expose format changes**~~ - ✅ Implemented: `FormatChanged` revision type + `FormatChange` details (`DetectFormatChanges`, native `w:rPrChange`)
4. ~~**Generate native move markup**~~ - ✅ Implemented: `Compare` emits native `w:moveFrom`/`w:moveTo` markup when `DetectMoves` is on
5. ~~**Granular format change details**~~ - ✅ Implemented: `FormatChange.ChangedPropertyNames` + old/new property dictionaries

### Addressed structurally by the IR diff engine (`DocxDiff`)

- **Add revision context / position information** — `DocxDiffRevision` carries `LeftAnchor`/`RightAnchor` (`kind:scope:unid`), locating each revision in the document model and interoperating with `DocxSession`. See [`ir_diff_engine.md`](./ir_diff_engine.md).
- **Diff-as-data** — `DocxDiff.GetEditScriptJson` serializes the edit script for storage/transport/audit.
