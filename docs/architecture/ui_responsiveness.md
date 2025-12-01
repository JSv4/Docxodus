# UI Responsiveness Architecture

## Problem Statement (Issue #44)

The Docxodus WASM runtime executes synchronously on the browser's main thread. When processing large documents, this blocks the event loop, preventing:
- React state updates from rendering (loading spinners don't appear)
- User interactions from being processed
- Animations from running smoothly

Even though API functions like `convertDocxToHtml()` are `async`, the actual WASM call inside them is synchronous and blocks until complete.

## Solution Overview

A three-phase approach addresses this problem with increasing sophistication:

| Phase | Approach | Blocking | Complexity | Status |
|-------|----------|----------|------------|--------|
| 1 | Frame Yielding | Yes (after initial paint) | Low | **Implemented** |
| 2 | Web Worker | No | Medium | Planned |
| 3 | Lazy Loading | No (per-page) | High | Planned |

## Phase 1: Frame Yielding (Implemented)

### Mechanism

Before every heavy WASM operation, the npm wrapper yields to the browser using the double-`requestAnimationFrame` pattern:

```typescript
async function yieldToMain(): Promise<void> {
  if (typeof requestAnimationFrame === "undefined") {
    return; // Skip in Node.js/SSR
  }

  await new Promise<void>((resolve) => {
    requestAnimationFrame(() => {
      requestAnimationFrame(() => resolve());
    });
  });
}
```

### Why Double-rAF?

1. **First rAF**: Schedules callback for the next animation frame
2. **Second rAF**: Ensures the first frame actually painted before continuing

This guarantees that any pending DOM updates (like showing a loading spinner) are committed and painted before the blocking WASM work begins.

### Functions with Yielding

- `convertDocxToHtml()` - Heavy HTML conversion
- `compareDocuments()` - Document comparison
- `compareDocumentsToHtml()` - Comparison + conversion
- `getRevisions()` - Revision extraction
- `addAnnotation()` - Document modification
- `addAnnotationWithTarget()` - Document modification
- `getDocumentStructure()` - Structure analysis

### Usage

No API changes required. Yielding is automatic:

```typescript
// Before: Loading spinner didn't appear
setLoading(true);
const html = await convertDocxToHtml(doc); // Blocked immediately
setLoading(false);

// After: Loading spinner appears before conversion starts
setLoading(true);
const html = await convertDocxToHtml(doc); // Yields, then blocks
setLoading(false);
```

### Limitations

- UI still freezes during WASM execution (after initial paint)
- Long operations (10+ seconds) will still feel unresponsive
- No progress indication during conversion

## Phase 2: Web Worker (Planned)

### Architecture

```
Main Thread                           Web Worker
┌─────────────────────────┐          ┌─────────────────────────┐
│ React App               │          │ docxodus.worker.ts      │
│                         │          │                         │
│ ┌─────────────────────┐ │  post   │ ┌─────────────────────┐ │
│ │ worker-proxy.ts     │─┼────────▶│ │ WASM Runtime        │ │
│ │ - Manages worker    │ │ Message │ │ - DocumentConverter │ │
│ │ - Handles responses │◀┼─────────│ │ - DocumentComparer  │ │
│ └─────────────────────┘ │          │ └─────────────────────┘ │
└─────────────────────────┘          └─────────────────────────┘
```

### Key Design Decisions

1. **Separate WASM instance**: Worker loads its own dotnet.js runtime
2. **Transferable bytes**: Document `Uint8Array` is transferred (zero-copy)
3. **Streaming-ready API**: Message structure supports future chunked output

### API Design (Preliminary)

```typescript
// Opt-in worker-based conversion
import { createWorkerDocxodus } from 'docxodus/worker';

const docxodus = await createWorkerDocxodus();
const html = await docxodus.convertDocxToHtml(doc); // Non-blocking!
```

### Benefits

- Main thread remains responsive during entire operation
- UI animations continue smoothly
- User can interact with other parts of the application

## Phase 3: Lazy Loading (Planned)

### Concept

Instead of converting the entire document at once, generate content on-demand as the user scrolls:

```
1. User drops DOCX file

2. Worker: getDocumentMetadata() [fast, ~100ms]
   → Returns: page count, dimensions, section info

3. Main thread: Create placeholder pages
   → User immediately sees scrollable page list

4. As user scrolls to page N:
   Worker: renderPage(N)
   → Returns: HTML for just that page
```

### Required Changes

**New WmlToHtmlConverter.cs methods:**

```csharp
// Fast metadata extraction (no full render)
public static DocumentMetadata GetDocumentMetadata(
    WordprocessingDocument wordDoc,
    WmlToHtmlConverterSettings settings);

// Render specific content range
public static XElement RenderContentRange(
    WordprocessingDocument wordDoc,
    WmlToHtmlConverterSettings settings,
    int sectionIndex,
    int startParagraph,
    int endParagraph);
```

**DocumentMetadata structure:**

```csharp
public class DocumentMetadata {
    public List<SectionMetadata> Sections { get; set; }
    public int TotalParagraphs { get; set; }
    public int EstimatedPageCount { get; set; }
}

public class SectionMetadata {
    public double PageWidthPt { get; set; }
    public double PageHeightPt { get; set; }
    public double ContentHeightPt { get; set; }
    public int ParagraphCount { get; set; }
    public bool HasHeader { get; set; }
    public bool HasFooter { get; set; }
}
```

### Compatibility with pagination.ts

The current `PaginationEngine` already has:
- Pre-measured header/footer heights
- Single-pass forward-only algorithm
- Page containers that can accept content incrementally

The key change: instead of `paginate()` processing all blocks upfront, `paginatePage(pageIndex)` would request content from the worker on demand.

### Benefits

- Fast initial render (< 1 second for any document)
- Low memory usage (only visible pages in DOM)
- Smooth scrolling experience

## Testing

### Phase 1 Tests (npm/tests/docxodus.spec.ts)

```typescript
test.describe('Frame Yielding Tests (Issue #44)', () => {
  test('loading state is observable before conversion completes');
  test('multiple async operations yield properly');
  test('comparison operation yields to allow loading state');
  test('getRevisions yields before processing');
  test('annotation operations yield properly');
});
```

### Phase 2 Tests (Planned)

- Worker initialization
- Message passing roundtrip
- Large document conversion without main thread blocking
- Error handling across worker boundary

### Phase 3 Tests (Planned)

- Metadata extraction speed
- Page rendering accuracy
- Memory usage under scroll stress
- Content integrity across lazy-loaded pages

## Migration Path

| Current Code | Phase 1 (Now) | Phase 2 | Phase 3 |
|-------------|---------------|---------|---------|
| `convertDocxToHtml()` | No change | `workerDocxodus.convertDocxToHtml()` | `workerDocxodus.convertDocxToHtml({ lazy: true })` |
| Manual loading state | Works correctly | Works correctly | Automatic per-page loading |

All phases are backwards-compatible. Consumers can adopt new features incrementally.
