# WASM Performance Profiling Results

This document captures performance profiling results for the Docxodus WASM pagination system, conducted December 2025.

## Test Environment

- **Platform**: Linux (Ubuntu)
- **Browser**: Chromium (Playwright)
- **WASM Runtime**: .NET 8.0 Blazor WASM

## Summary of Findings

### 1. WASM Initialization Overhead

| Metric | Time |
|--------|------|
| Average initialization | 300-600ms |
| Min observed | 287ms |
| Max observed | 653ms |

This is a one-time cost per page load. Caching/reusing the WASM module eliminates this for subsequent operations.

### 2. GetDocumentMetadata Performance

`GetDocumentMetadata` was designed to be a lightweight operation for virtual scrolling, but profiling reveals it's unexpectedly slow:

| Document | Size | Paragraphs | GetDocumentMetadata | % of Total |
|----------|------|------------|---------------------|------------|
| RPR-FivePageTestDoc | 1.5 KB | 125 | 288ms | 33.6% |
| RPR-TenPageTestDoc | 2.0 KB | 250 | 305ms | 26.8% |
| HC007-Test-02 | 26.6 KB | 210 | 617ms | 9.0% |
| DB0016-DocDefaultStyles | 2.7 MB | 1 | 608ms | 12.1% |

**Root Cause**: The metadata extraction parses the entire document structure even though it doesn't render HTML. This parsing is then repeated when calling `RenderPageRange`.

### 3. Full Document Conversion Timing

#### Small/Medium Documents (HC007-Test-02, 26.6 KB, 210 paragraphs, 8 tables)

| Operation | Time | % of Total |
|-----------|------|------------|
| ConvertDocxToHtml | 2.37s | 34.7% |
| ConvertDocxToHtmlWithPagination | 1.65s | 24.2% |
| ConvertDocxToHtmlComplete | 1.49s | 21.8% |
| GetDocumentMetadata | 617ms | 9.0% |
| RenderPageRange (5 pages) | 355ms | 5.2% |
| RenderPageRange (1 page) | 316ms | 4.6% |
| Block measurement (client) | 29ms | 0.4% |
| DOM insertion | 2.3ms | 0.0% |
| **Total** | **6.82s** | |

#### Large Document (DB0016-DocDefaultStyles, 2.7 MB)

| Operation | Time | % of Total |
|-----------|------|------------|
| ConvertDocxToHtml | 1.73s | 34.3% |
| ConvertDocxToHtmlComplete | 1.16s | 23.1% |
| ConvertDocxToHtmlWithPagination | 1.16s | 23.0% |
| GetDocumentMetadata | 608ms | 12.1% |
| RenderPageRange (1 page) | 366ms | 7.3% |
| Block measurement (client) | 12ms | 0.2% |
| DOM insertion | 0.4ms | 0.0% |
| **Total** | **5.04s** | |

### 4. Page Range Rendering vs Full Render

```
Document: HC007-Test-02.docx (8 estimated pages)
═══════════════════════════════════════════════════
Metadata extraction: 541ms
Full render:         2.47s
Page range (1-2):    417ms
Speedup:             5.9x
═══════════════════════════════════════════════════
```

**Important**: The 5.9x speedup is for the render step only. When including the required `GetDocumentMetadata` call:
- Page range approach: 541ms + 417ms = **958ms**
- Full render: **2.47s**
- Actual speedup: **~2.6x**

### 5. Page Range Scaling Behavior

| Pages Rendered | Time | Notes |
|----------------|------|-------|
| 1 page | 758ms | Includes document parsing overhead |
| 2 pages | 438ms | Faster due to cached parsing |
| 3 pages | 381ms | |
| 5 pages | 352ms | |
| 8 pages | 399ms | |

**Observation**: First page render includes significant startup overhead. The per-page marginal cost is low after initial parsing.

### 6. Client-Side Operations (Nearly Free)

| Operation | Time |
|-----------|------|
| DOM insertion | 0.3-2.3ms |
| Block measurement (total) | 5-30ms |
| Average per block | 1.7ms |

Client-side pagination (measuring and flowing blocks) is essentially negligible compared to WASM conversion.

## Bottleneck Analysis

### Time Distribution (Typical Document)

```
┌────────────────────────────────────────────────────────────────────┐
│  Operation                                    Time      % Total    │
├────────────────────────────────────────────────────────────────────┤
│  ████████████  WASM Init                      300-600ms  (1x)     │
│  ████████████████████  GetDocumentMetadata    300-600ms  25-35%   │
│  ████████████████████████████  ConvertToHtml  1-2.5s     35-50%   │
│  ░░ DOM + Client Measurement                  <50ms      <1%      │
└────────────────────────────────────────────────────────────────────┘
```

### Primary Bottleneck: Duplicate Document Parsing

The current flow parses the document multiple times:

```
GetDocumentMetadata() → Opens DOCX, parses XML, extracts metadata
RenderPageRange()     → Opens DOCX AGAIN, parses XML AGAIN, renders HTML
```

Each call independently:
1. Opens the DOCX ZIP archive
2. Parses `document.xml`
3. Parses `styles.xml`
4. Traverses the entire element tree
5. Extracts section properties

## Implemented Optimization: Document Session Caching

### Results (December 2025)

Document session caching has been implemented and achieves significant performance gains:

| Metric | Value |
|--------|-------|
| **Time Savings** | **70%** |
| **Speedup** | **3.33x** |
| OLD approach (direct API) | 772.80ms |
| NEW approach (session) | 231.90ms |

### Performance Comparison

```
════════════════════════════════════════════════════════════════
PERFORMANCE COMPARISON: Session vs Direct API
════════════════════════════════════════════════════════════════
OLD APPROACH (direct API, parses document twice):
  GetDocumentMetadata:     541.40ms (document parsed)
  RenderPageRange:         231.40ms (document parsed AGAIN)
  Total:                   772.80ms

NEW APPROACH (session-based, parses once):
  CreateDocumentSession:   200.50ms (document parsed + cached)
  GetSessionMetadata:        0.00ms (instant, from cache)
  RenderSessionPageRange:   31.40ms (uses cached document)
  Total:                   231.90ms

SAVINGS: 540.90ms (70%)
SPEEDUP: 3.33x
════════════════════════════════════════════════════════════════
```

### TypeScript API

The session API is available in the TypeScript package:

```typescript
import {
  createDocumentSession,
  renderSessionPageRange,
  getSessionMetadata,
  closeDocumentSession,
  getActiveSessions,
} from 'docxodus';

// Create a session (parses document once)
const session = await createDocumentSession(docxBytes);

// Metadata is immediately available (no extra parsing)
console.log(`Document has ${session.metadata.estimatedPageCount} pages`);

// Render pages efficiently (uses cached document)
const html1 = await renderSessionPageRange(session.sessionId, 1, 3);
const html2 = await renderSessionPageRange(session.sessionId, 4, 6);

// Clean up when done
await closeDocumentSession(session.sessionId);
```

### WASM Implementation

The session caching is implemented in `DocumentSession.cs`:

- `CreateSession(bytes)` - Parse document once, cache in memory, return session ID + metadata
- `GetSessionMetadata(sessionId)` - Return cached metadata (instant)
- `RenderSessionPageRange(sessionId, start, end, ...)` - Render using cached document
- `RenderSessionDocument(sessionId, ...)` - Full render using cached document
- `CloseSession(sessionId)` - Release cached document from memory
- `GetActiveSessions()` - List all active sessions (for debugging/monitoring)

### Decision Guide

| Document Size | Recommended Approach |
|---------------|---------------------|
| < 50 pages | Full render (simpler, overhead not worth it) |
| 50-200 pages | Page range with session caching |
| > 200 pages | Virtual scroll with aggressive caching |

### Future Optimizations

**Background Pre-Parsing**:
```typescript
// After file upload, start parsing in background
const session = await createDocumentSession(bytes);

// When user views document (parsing already done)
const metadata = session.metadata;  // instant, already available
const html = await renderSessionPageRange(session.sessionId, 1, 3);
```

## Running the Profiling Tests

```bash
cd npm
npm run pretest
npx playwright test profiling.spec.ts --project=chromium --reporter=list
```

The profiling harness is also available interactively at:
```
npm run pretest
python3 -m http.server 8082 --directory dist/wasm
# Open http://localhost:8082/profiling-harness.html
```

## Test Files Used

| File | Size | Paragraphs | Tables | Est. Pages |
|------|------|------------|--------|------------|
| RPR-FivePageTestDoc.docx | 1.5 KB | 125 | 0 | 5 |
| RPR-TenPageTestDoc.docx | 2.0 KB | 250 | 0 | 10 |
| HC007-Test-02.docx | 26.6 KB | 210 | 8 | 8 |
| DB0016-DocDefaultStyles.docx | 2.7 MB | 1 | 0 | 1 |

## Related Documentation

- [Pagination Architecture](pagination.md) - Overall pagination system design
- [Virtual Scrolling](pagination.md#virtual-scrolling--lazy-loading-issue-31) - Lazy loading implementation
