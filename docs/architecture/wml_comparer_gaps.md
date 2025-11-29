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

## 2. Move Detection Not Implemented

**Status:** Gap

The comparison algorithm does not attempt to detect when content has been moved from one location to another. Instead, moved content appears as:
1. A deletion at the original location
2. An insertion at the new location

Word's native comparison feature can detect moves and mark them with `w:moveFrom` / `w:moveTo` elements, which the HTML renderer can then display distinctly from regular insertions/deletions.

### Recommendation

Implement move detection by:
1. After identifying deletions and insertions, compare their text content
2. If a deletion's text closely matches an insertion's text (above a similarity threshold), mark them as a move pair
3. Add `Moved` to `WmlComparerRevisionType` and track source/destination locations

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

### High Priority

1. **Add move detection** - Distinguish moved content from delete+insert pairs
2. **Expose format changes** - Add `FormatChange` revision type for formatting-only modifications

### Medium Priority

3. **Add revision context** - Include paragraph number or surrounding text for better UX
4. **Link related revisions** - Connect move pairs and other related changes

### Low Priority

5. **Position information** - Add character/word offsets for precise location
6. **Granular format change details** - Specify exactly which properties changed
