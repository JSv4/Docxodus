# Pagination Architecture

This document describes the client-side pagination system that provides a PDF.js-style paginated view of converted documents.

## Overview

The pagination system works in two phases:
1. **Server/WASM side** (C#): Generate HTML with pagination metadata and CSS
2. **Client side** (TypeScript): Measure content and flow it into fixed-size page containers

This separation allows the computationally expensive document conversion to happen once, while the layout-dependent pagination runs in the browser where accurate content measurement is possible.

## Architecture Diagram

```
┌─────────────────────────────────────────────────────────────────┐
│                    Document Conversion (C#)                      │
│  WmlToHtmlConverter.ConvertToHtml()                             │
│  - Extracts page dimensions from w:sectPr                       │
│  - Generates HTML with data attributes for page metadata        │
│  - Outputs staging structure for client-side pagination         │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│                      HTML Output Structure                       │
│  <div id="pagination-staging" class="page-staging">             │
│    <div data-section-index="0"                                  │
│         data-page-width="612" data-page-height="792"            │
│         data-content-width="468" data-content-height="648"      │
│         data-margin-top="72" data-margin-left="72" ...>         │
│      <!-- Document content here -->                             │
│    </div>                                                       │
│  </div>                                                         │
│  <div id="pagination-container" class="page-container"></div>   │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│                  Client-Side Pagination (TS)                     │
│  PaginationEngine.paginate()                                    │
│  1. Parse dimensions from data attributes                       │
│  2. Set staging width for accurate line wrapping                │
│  3. Measure block heights with getBoundingClientRect()          │
│  4. Flow blocks into pages based on available height            │
│  5. Clone content into page containers                          │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│                    Rendered Page Structure                       │
│  <div class="page-container">                                   │
│    <div class="page-box" style="width:612pt; height:792pt;      │
│         transform:scale(0.8); overflow:hidden">                 │
│      <div class="page-content" style="position:absolute;        │
│           top:72pt; left:72pt; width:468pt; height:648pt">      │
│        <!-- Cloned content blocks -->                           │
│      </div>                                                     │
│      <div class="page-number">1</div>                           │
│    </div>                                                       │
│  </div>                                                         │
└─────────────────────────────────────────────────────────────────┘
```

## Key Components

### WmlToHtmlConverterSettings (C#)

```csharp
// Pagination-related settings
public PaginationMode RenderPagination;        // None or Paginated
public double PaginationScale;                  // Scale factor (1.0 = 100%)
public string PaginationCssClassPrefix;        // CSS class prefix (default: "page-")
```

### PageDimensions (TypeScript)

```typescript
interface PageDimensions {
  pageWidth: number;      // Full page width in points
  pageHeight: number;     // Full page height in points
  contentWidth: number;   // Content area width (page minus margins)
  contentHeight: number;  // Content area height (page minus margins)
  marginTop: number;      // Top margin in points
  marginRight: number;    // Right margin in points
  marginBottom: number;   // Bottom margin in points
  marginLeft: number;     // Left margin in points
}
```

### PaginationEngine Class

The main pagination engine handles the content flow:

```typescript
class PaginationEngine {
  constructor(staging: HTMLElement | string,
              container: HTMLElement | string,
              options?: PaginationOptions)

  paginate(): PaginationResult
}

interface PaginationOptions {
  scale?: number;           // Zoom level (default: 1)
  cssPrefix?: string;       // CSS class prefix (default: "page-")
  showPageNumbers?: boolean; // Show page numbers (default: true)
  pageGap?: number;         // Gap between pages in pixels (default: 20)
}
```

## Content Flow Algorithm

The pagination algorithm in `flowToPages()` handles several cases:

1. **Normal blocks**: Add to current page if height fits, otherwise start new page
2. **Page breaks**: Explicit breaks (via CSS or data attribute) force a new page
3. **Keep with next**: Blocks marked with `data-keep-with-next="true"` consider the next block's height
4. **Oversized blocks**: Blocks taller than a page get their own page (overflow is clipped)

### Block Measurement

Each block is measured with `getBoundingClientRect()` for content dimensions, plus `getComputedStyle()` for margins:

```typescript
interface MeasuredBlock {
  element: HTMLElement;
  heightPt: number;      // Content + padding + border (excluding margins)
  marginTopPt: number;   // Top margin
  marginBottomPt: number; // Bottom margin
  // ... other properties
}
```

### Margin Collapsing

CSS vertical margins collapse between adjacent blocks. The algorithm accounts for this:

```typescript
// Track the previous block's bottom margin
let prevMarginBottomPt = 0;

for (const block of blocks) {
  // Calculate effective margin gap (collapsed)
  const isFirstOnPage = currentContent.length === 0;
  let effectiveMarginTop = block.marginTopPt;
  if (!isFirstOnPage) {
    // Margin collapsing: gap is max(prevBottom, currTop), not sum
    effectiveMarginTop = Math.max(block.marginTopPt, prevMarginBottomPt) - prevMarginBottomPt;
  }

  // Total space = effective margin + content + bottom margin
  const blockSpace = effectiveMarginTop + block.heightPt + block.marginBottomPt;

  if (blockSpace <= remainingHeight) {
    currentContent.push(block.element.cloneNode(true));
    remainingHeight -= blockSpace;
    prevMarginBottomPt = block.marginBottomPt;
  } else {
    // Start new page...
  }
}
```

This ensures accurate pagination that matches browser rendering behavior.

## Scaling Implementation

Scaling uses a hybrid approach with CSS `zoom` (preferred) and `transform` (fallback):

1. Page box is created at **full document size** (e.g., 612pt × 792pt)
2. Content area is positioned at **full margins** with **full dimensions**
3. **CSS `zoom`** is applied first (better text rendering, affects layout naturally)
4. **CSS `transform: scale()`** is also set as fallback for browsers without zoom support
5. `transform-origin: top left` keeps the top-left corner fixed
6. Negative margins compensate for transform not affecting layout (only needed when zoom unsupported)

### Why Zoom + Transform?

- **CSS `zoom`**: Non-standard but widely supported (Chrome, Safari, Edge, Firefox 126+). Affects layout directly, renders text at target resolution for crisp output.
- **CSS `transform`**: Standard but doesn't affect layout. Text may appear blurry at fractional scales because it's rasterized at original size then scaled.

When both are set, browsers that support `zoom` use it and ignore the redundant transform scaling effect. Browsers without zoom support fall back to transform.

### Performance Optimizations

- **`will-change: transform`**: Hints browser to create a GPU compositing layer
- **`contain: layout paint`**: Isolates the page box for layout and paint operations, preventing changes from affecting siblings

This approach ensures:
- Content is properly contained and clipped by `overflow: hidden`
- Scaling is uniform for all page elements
- Crisp text rendering in modern browsers
- Fallback for older browsers

## CSS Structure

The pagination system generates CSS classes:

```css
/* Staging area - hidden for measurement */
.page-staging {
  position: absolute;
  left: -9999px;
  visibility: hidden;
}

/* Container with dark background (PDF.js style) */
.page-container {
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 20px;
  padding: 20px;
  background: #525659;
  min-height: 100vh;
}

/* Individual page box */
.page-box {
  background: white;
  box-shadow: 0 2px 8px rgba(0,0,0,0.3);
  position: relative;
  overflow: hidden;
  box-sizing: border-box;
}

/* Content area within page */
.page-content {
  position: absolute;
  overflow: hidden;
  transform-origin: top left;
}

/* Page number indicator */
.page-number {
  position: absolute;
  bottom: 8px;
  width: 100%;
  text-align: center;
  font-size: 11px;
  color: #666;
}
```

## React Integration

### usePagination Hook

```tsx
function Viewer({ html }: { html: string }) {
  const containerRef = useRef<HTMLDivElement>(null);
  const { result, isPaginating, error } = usePagination(html, containerRef, {
    scale: 0.8,
    showPageNumbers: true
  });

  return (
    <div ref={containerRef}>
      {result && <span>Total pages: {result.totalPages}</span>}
    </div>
  );
}
```

### PaginatedDocument Component

```tsx
import { PaginatedDocument, PaginationMode } from 'docxodus/react';

function Viewer() {
  const { isReady, convertToHtml } = useDocxodus();
  const [html, setHtml] = useState<string | null>(null);

  const handleFile = async (file: File) => {
    const result = await convertToHtml(file, {
      paginationMode: PaginationMode.Paginated
    });
    setHtml(result);
  };

  return html ? (
    <PaginatedDocument
      html={html}
      scale={0.8}
      showPageNumbers={true}
      onPageVisible={(pageNum) => console.log(`Page ${pageNum} visible`)}
    />
  ) : null;
}
```

## Unit Conversion

The system works with points (pt) throughout for consistency with Word documents:

```typescript
// CSS points to pixels (96 DPI screen assumption)
function ptToPx(pt: number): number {
  return pt / 0.75; // 96 / 72 = 1.333...
}

// Pixels to points
function pxToPt(px: number): number {
  return px * 0.75; // 72 / 96 = 0.75
}
```

Important: CSS `pt` units render at 96/72 ratio to pixels, so setting `width: 100pt` renders as ~133.33px.

## Limitations

1. **No block splitting**: Blocks taller than a page are not split; they overflow (clipped)
2. **Paragraph-level granularity**: Pagination operates on block elements, not individual lines
3. **Limited widow/orphan control**: `keepWithNext` is supported, but line-level control is not
4. **No column support**: Multi-column layouts are flattened to single column
5. **Browser-dependent measurement**: Content measurement depends on browser rendering

## Related Documentation

- [Paginated Headers and Footers](paginated_headers_footers.md) - How headers and footers are rendered within paginated pages

## Future Enhancements

1. **Block splitting**: Split oversized paragraphs/tables across pages
2. **Line-level pagination**: Measure and flow individual lines for better control
3. **Column support**: Handle multi-column sections
4. **Virtual scrolling**: Only render visible pages for large documents
5. **Server-side rendering**: Pre-compute pagination for static documents
