# WASM Performance Optimization Plan

**Created:** December 2025
**Status:** All Phases Complete (Phase 3 added ~50% speedup)

## Executive Summary

Profiling revealed that **virtualization does not improve performance** for document rendering. The bottleneck is WASM conversion itself, not DOM operations. This plan outlines:

1. Removal of unused virtualization code
2. Targeted optimizations based on profiling data
3. Simplified architecture recommendations

## Profiling Results

### Time Distribution (8-page document with tables)

```
Phase                               Time      %
═══════════════════════════════════════════════════════════════════════════════
Transform_1_ConvertToHtml         1.67s   55.7%  ← Contains duplicate preprocessing
Style_1_FormattingAssembler     673.60ms  22.5%  ← Style resolution (OPTIMIZATION TARGET)
Preprocess_1_AcceptRevisions    271.90ms   9.1%  ← Revision processing (OPTIMIZATION TARGET)
Preprocess_2_SimplifyMarkup     187.60ms   6.3%
Load_3_GetXDocument             142.40ms   4.8%
Load_2_OpenXml_Open              25.60ms   0.9%  ← ZIP decompression (NOT a bottleneck)
Serialize_1_ToString             19.70ms   0.7%
═══════════════════════════════════════════════════════════════════════════════
TOTAL                              3.00s
```

### Key Finding: Virtualization Doesn't Help

| Approach | Initial Load | After 1 Scroll | Winner |
|----------|--------------|----------------|--------|
| One-shot (render all) | 70ms | N/A | **Winner for simple docs** |
| Virtualized | 65ms | 113ms (+60%) | Loses after scrolling |

**Root cause:** Each virtualized page render re-parses the document. The overhead of multiple parse operations exceeds the savings from rendering less content.

## Phase 1: Remove Unused Virtualization Code ✅ COMPLETE

### Completed Changes

1. **`npm/src/types.ts`** ✅
   - Removed `RenderPageRangeOptions` interface
   - Removed `VirtualPaginationOptions`, `VirtualPaginationResult`, and related types
   - Removed `PageLayout`, `PageLayoutDimensions`, `CachedBlock`, `SearchResult`, `SearchOptions`
   - Removed worker page-range request/response types

2. **`npm/src/index.ts`** ✅
   - Removed `renderPageRange()` function
   - Removed `renderSessionPageRange()` function
   - Kept `createDocumentSession()`, `renderSessionDocument()`, `closeDocumentSession()` (caching still valuable)

3. **`npm/src/react.ts`** ✅
   - Removed `useVirtualPagination` hook
   - Removed `VirtualPaginatedDocument` component
   - Removed `VirtualPaginationOptions`, `UseVirtualPaginationResult` interfaces

4. **`npm/src/worker-proxy.ts`** ✅
   - Removed `renderPageRange()` method

5. **`npm/src/docxodus.worker.ts`** ✅
   - Removed `handleRenderPageRange()` function
   - Removed page-range case handler

6. **`wasm/DocxodusWasm/DocumentConverter.cs`** ✅
   - Removed `RenderPageRange()` method
   - Removed `RenderPageRangeFull()` method
   - Kept `GetDocumentMetadata()` (useful for document info)

7. **`wasm/DocxodusWasm/DocumentSession.cs`** ✅
   - Removed `RenderSessionPageRange()` method
   - Kept session caching infrastructure (still provides 3x speedup)

### Future Cleanup (Optional)

These files still contain RenderPageRange code but it's now unreachable from the public API:

- **`Docxodus/WmlToHtmlConverter.cs`** - `RenderPageRange()` region (~770 lines) can be removed later
- **`Docxodus.Tests/DocumentMetadataTests.cs`** - `RenderPageRange` tests (~480 lines) can be removed later

## Phase 2: Optimize AcceptRevisions ✅ COMPLETE

### Implementation (WmlToHtmlConverter.cs:647-762)

```csharp
// Only accept revisions if NOT rendering tracked changes AND document has tracked changes
// This optimization saves ~9% of conversion time for documents without revisions
if (!htmlConverterSettings.RenderTrackedChanges && HasTrackedChanges(wordDoc))
{
    RevisionAccepter.AcceptRevisions(wordDoc);
}

private static bool HasTrackedChanges(WordprocessingDocument wordDoc)
{
    var mainPart = wordDoc.MainDocumentPart;
    if (mainPart == null) return false;

    var xDoc = mainPart.GetXDocument();
    var body = xDoc.Root?.Element(W.body);
    if (body == null) return false;

    // Check for any revision elements - this is a fast descendant scan
    return body.Descendants().Any(e =>
        e.Name == W.ins || e.Name == W.del ||
        e.Name == W.moveFrom || e.Name == W.moveTo ||
        e.Name == W.rPrChange || e.Name == W.pPrChange ||
        e.Name == W.tblPrChange || e.Name == W.tcPrChange ||
        e.Name == W.sectPrChange);
}
```

**Expected savings:** ~270ms (9%) on documents without tracked changes

## Phase 3: Optimize FormattingAssembler ✅ COMPLETE

### Analysis

The FormattingAssembler resolves style inheritance for each paragraph and run. Key methods:
- `ParagraphStyleRollup()` - Walks up paragraph style hierarchy
- `ParaStyleParaPropsStack()` - Builds style property stack
- `AnnotateRuns()` - Similar for run styles

### Complexity Factors Addressed

1. **Style chain depends on list item info** - Solved by caching only non-list-item paragraphs; list items are computed per-paragraph
2. **Table context affects styles** - Table styles are merged separately (unchanged)
3. **Toggle properties need special handling** - Handled correctly by existing MergeStyleElement

### Implemented Optimizations

1. **Pre-indexed style lookups** - Build dictionary indexes once at start, O(1) lookups everywhere
2. **Paragraph style rollup caching** - Cache rolled-up properties by style name for non-list paragraphs

### Results

**Actual effort:** ~1 hour for implementation + testing
**Actual savings:** ~1.2s (50-55%) on 8-page document - **far exceeded expectations**

## Phase 4: Recommended Architecture ✅ COMPLETE

### Simplified Flow

```
Document Render:
  1. convertDocxToHtml(bytes, options)  // Single call, no session management
     - Full HTML render
     - Use CSS content-visibility: auto for browser-native optimization

Optional (for info only):
  2. getDocumentMetadata(bytes)         // If you need page count, etc.
```

### Client-Side Recommendation

```typescript
// Simple, fast, minimal API
const html = await convertDocxToHtml(docxFile, {
  paginationMode: PaginationMode.Paginated,
  renderFootnotesAndEndnotes: true
});

// Let browser handle "virtualization" natively
container.style.contentVisibility = 'auto';
container.style.containIntrinsicSize = '0 50000px';
container.innerHTML = html;
```

## Implementation Order

1. **Remove virtualization code** (Phase 1) ✅ - Reduced complexity
2. **AcceptRevisions optimization** (Phase 2) ✅ - ~9% improvement on clean documents
3. **Remove session API** (Phase 4) ✅ - Simplified API surface
4. **FormattingAssembler caching** (Phase 3) ✅ - **~50-55% improvement** (December 2025)

## Phase 3: FormattingAssembler Caching ✅ COMPLETE

### Implementation Details

Two key optimizations were implemented in `FormattingAssembler.cs`:

#### Optimization #1: Pre-Indexed Style Lookups

Instead of O(n) linear searches through all styles at each inheritance level, styles are now indexed once at the start of formatting assembly:

```csharp
// FormattingAssemblerInfo now includes:
public Dictionary<string, XElement> ParagraphStyleIndex;
public Dictionary<string, XElement> CharacterStyleIndex;
public Dictionary<string, XElement> AllStylesIndex;

// IndexStylesDocument() builds these once, then all lookups are O(1)
```

**Files changed:**
- `FormattingAssembler.cs:109-136` - `IndexStylesDocument()` method
- `FormattingAssembler.cs:2375-2387` - `ParaStyleParaPropsStack()` uses indexed lookup
- `FormattingAssembler.cs:2708-2722` - `ParaStyleRunPropsStack()` uses indexed lookup
- `FormattingAssembler.cs:2754-2786` - `CharStyleStack()` uses indexed lookup

#### Optimization #2: Paragraph Style Rollup Caching

For non-list-item paragraphs, the rolled-up style properties are cached by style name:

```csharp
// FormattingAssemblerInfo now includes:
public Dictionary<string, XElement> CachedParagraphStyleRollups;

// Cache hit returns cloned element; cache miss computes and stores
```

**Files changed:**
- `FormattingAssembler.cs:2339-2359` - `ParagraphStyleRollupInternal()` with caching

### Performance Results (HC007-Test-02.docx - 210 paragraphs, 8 tables)

| Operation | Before | After | Improvement |
|-----------|--------|-------|-------------|
| GetDocumentMetadata | 617ms | 278ms | **55% faster** |
| ConvertDocxToHtml | 2.37s | 1.15s | **51% faster** |
| ConvertDocxToHtmlWithPagination | 1.65s | 752ms | **54% faster** |
| ConvertDocxToHtmlComplete | 1.49s | 667ms | **55% faster** |

## Success Metrics

| Metric | Before | After Phase 2 | After Phase 3 |
|--------|--------|---------------|---------------|
| 8-page doc conversion | 3.00s | ~2.7s | **~1.15s** |
| Code complexity | High | Medium | Low |
| API surface | Large | Small | Small |

## Related Documents

- [Profiling Results](profiling-results.md) - Detailed timing data
