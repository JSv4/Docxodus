# WmlComparer.cs - Gaps and Deficiencies

This document catalogs known gaps, limitations, and areas for improvement in the WmlComparer (document comparison engine).

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

**Note:** This implementation does NOT generate Word-native `w:moveFrom`/`w:moveTo` markup in the document. It only affects the `GetRevisions()` API. Generating native move markup would require deeper integration into the comparison algorithm and is a potential future enhancement.

## 3. Format Change Detection Not Exposed

**Status:** Gap

When formatting changes occur without text changes, the comparison engine may not surface these as distinct revisions. Word documents can track formatting changes via:
- `w:rPrChange` - Run property changes (font, size, bold, etc.)
- `w:pPrChange` - Paragraph property changes (alignment, spacing, etc.)
- `w:sectPrChange` - Section property changes
- `w:tblPrChange` - Table property changes

### Recommendation

Add a `FormatChange` revision type that captures:
- The element affected
- The old formatting properties
- The new formatting properties

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

### High Priority

1. **Expose format changes** - Add `FormatChange` revision type for formatting-only modifications

### Medium Priority

2. **Add revision context** - Include paragraph number or surrounding text for better UX
3. **Generate native move markup** - Produce `w:moveFrom`/`w:moveTo` elements in comparison output

### Low Priority

4. **Position information** - Add character/word offsets for precise location
5. **Granular format change details** - Specify exactly which properties changed
