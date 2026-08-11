# Paginated Headers and Footers Architecture

This document describes the architecture for rendering document headers and footers within the client-side pagination system.

**Source Files:**
- `Docxodus/WmlToHtmlConverter.cs` - C# header/footer rendering and registry generation
- `npm/src/pagination.ts` - TypeScript pagination engine with header/footer cloning

## Overview

Word documents have headers and footers that repeat on each page, potentially with different content for:
- **First page** of a section (when "Different first page" is enabled)
- **Odd pages** (default header/footer)
- **Even pages** (when "Different odd & even pages" is enabled)

When pagination is enabled, headers and footers must appear on each rendered page, not as document-level elements.

## The Problem

Without this feature, enabling both `RenderHeadersAndFooters` and `RenderPagination` results in:
- Headers/footers rendered once at the document level
- Pagination creating page boxes without headers/footers
- Visual mismatch where header/footer appear outside the paginated view

```
┌──────────────────────────────────┐
│ <header class="document-header"> │  ← Rendered once, outside pages
└──────────────────────────────────┘
┌──────────────────────────────────┐
│     Page 1 (no header/footer)    │  ← Pages missing headers/footers
│          [content...]            │
└──────────────────────────────────┘
┌──────────────────────────────────┐
│     Page 2 (no header/footer)    │
│          [content...]            │
└──────────────────────────────────┘
┌──────────────────────────────────┐
│ <footer class="document-footer"> │  ← Rendered once, outside pages
└──────────────────────────────────┘
```

## Solution Architecture

### Strategy: Per-Page Header/Footer Cloning

Store header/footer content in a hidden "registry" within the pagination staging area. The TypeScript pagination engine clones the appropriate header/footer into each page during rendering.

```
┌─────────────────────────────────────────────────────────────────────┐
│                      C# HTML Generation                              │
│  - Detect pagination + headers/footers both enabled                 │
│  - Render headers/footers into hidden registry (not document-level) │
│  - Include header/footer heights in section data attributes         │
└─────────────────────────────────────────────────────────────────────┘
                                │
                                ▼
┌─────────────────────────────────────────────────────────────────────┐
│                      HTML Output Structure                           │
│  <div id="pagination-staging">                                      │
│    <!-- Hidden registry for header/footer templates -->             │
│    <div id="pagination-hf-registry" style="display:none">           │
│      <div data-section="0" data-hf-type="header-default">...</div>  │
│      <div data-section="0" data-hf-type="footer-default">...</div>  │
│      <div data-section="0" data-hf-type="header-first">...</div>    │
│    </div>                                                           │
│    <!-- Section content with dimension data -->                     │
│    <div data-section-index="0"                                      │
│         data-header-height="36" data-footer-height="36" ...>        │
│      [content]                                                      │
│    </div>                                                           │
│  </div>                                                             │
└─────────────────────────────────────────────────────────────────────┘
                                │
                                ▼
┌─────────────────────────────────────────────────────────────────────┐
│                  TypeScript Pagination Engine                        │
│  1. Parse header/footer registry                                    │
│  2. For each page created:                                          │
│     - Determine which header/footer applies (first vs default)      │
│     - Clone header into page's top margin area                      │
│     - Clone footer into page's bottom margin area                   │
└─────────────────────────────────────────────────────────────────────┘
                                │
                                ▼
┌─────────────────────────────────────────────────────────────────────┐
│                    Rendered Page Structure                           │
│  <div class="page-box">                                             │
│    <div class="page-header" style="top:36pt;">                      │
│      [cloned header content]                                        │
│    </div>                                                           │
│    <div class="page-content" style="top:72pt; height:648pt;">       │
│      [flowed content]                                               │
│    </div>                                                           │
│    <div class="page-footer" style="bottom:36pt;">                   │
│      [cloned footer content]                                        │
│    </div>                                                           │
│  </div>                                                             │
└─────────────────────────────────────────────────────────────────────┘
```

## Word Document Header/Footer Model

### Section Properties (w:sectPr)

Each section can define its own headers and footers:

```xml
<w:sectPr>
  <w:headerReference w:type="default" r:id="rId6"/>
  <w:headerReference w:type="first" r:id="rId7"/>
  <w:headerReference w:type="even" r:id="rId8"/>
  <w:footerReference w:type="default" r:id="rId9"/>
  <w:footerReference w:type="first" r:id="rId10"/>
  <w:titlePg/>  <!-- Enables "Different first page" -->
  <w:pgMar w:header="720" w:footer="720" .../>
</w:sectPr>
```

### Header/Footer Types

| Type | w:type Value | When Used |
|------|--------------|-----------|
| Default | `"default"` | Odd pages, or all pages if no even header |
| First | `"first"` | First page of section (requires `w:titlePg`) |
| Even | `"even"` | Even pages (requires "Different odd & even" setting) |

### Reference Inheritance Across Sections

A `w:sectPr` that omits a reference type does **not** have an empty story of that
type: each story type is linked to the preceding section until that type is
explicitly replaced. Only the first section starts with blank stories.

```xml
<!-- Section 1: defines default + first -->
<w:sectPr>
  <w:headerReference w:type="default" r:id="rId6"/>
  <w:headerReference w:type="first" r:id="rId7"/>
</w:sectPr>

<!-- Section 2: defines nothing — BOTH stories still render, inherited from section 1 -->
<w:sectPr>
  <w:pgSz w:w="12240" w:h="15840"/>
</w:sectPr>
```

`EffectiveHeaderFooterReferences` (in `WmlToHtmlConverter.cs`) carries the
effective reference per story type as sections are visited in document order,
and is the single owner of that rule for both the paginated registry and
`GetDocumentMetadata`. A section with no `w:sectPr` at all (`CollectSectionData`'s
documented `null` ⇒ defaults contract) records nothing and leaves the inherited
references untouched.

Inheritance is orthogonal to *visibility*: an inherited or explicit `first`/`even`
reference is still inert unless `w:titlePg` / the document-global
`w:evenAndOddHeaders` switch it on (see Limitations).

### Margin Model

Word's page margins include header/footer space:

```
┌─────────────────────────────────┐ ← Page top
│     ↕ w:header (header dist)    │
│  ┌─────────────────────────┐    │
│  │   HEADER CONTENT        │    │
│  └─────────────────────────┘    │
│     ↕ (space to margin top)     │
├─────────────────────────────────┤ ← w:top (margin top)
│                                 │
│        BODY CONTENT             │
│                                 │
├─────────────────────────────────┤ ← w:bottom (margin bottom)
│     ↕ (space to footer)         │
│  ┌─────────────────────────┐    │
│  │   FOOTER CONTENT        │    │
│  └─────────────────────────┘    │
│     ↕ w:footer (footer dist)    │
└─────────────────────────────────┘ ← Page bottom
```

- `w:header`: Distance from page top to header text top (in twips)
- `w:footer`: Distance from page bottom to footer text bottom (in twips)
- `w:top`: Distance from page top to body content top
- `w:bottom`: Distance from page bottom to body content bottom

These are **four independent distances from the paper edge**, not two nested boxes. The header is
laid out from `w:header` *downward*; the footer from `w:footer` *upward*. The body then starts at
`w:top` unless the header has already run past it, and ends at `w:bottom` unless the footer has
already climbed above it:

```
bodyTop    = max(w:top,    w:header + headerContentHeight)
bodyBottom = min(pageHeight - w:bottom, pageHeight - w:footer - footerContentHeight)
```

A section with no story of that kind contributes nothing: the distance alone reserves no space,
so a header-less document keeps its body on `w:top`. `resolvePageBands()` in
`npm/src/page-geometry.ts` is the single owner of these four lines, and
`PaginationEngine.getPageBands()` feeds it the registry's measured story heights.

Anchoring the bands to the MARGINS instead — header bottom-aligned above `w:top`, footer
top-aligned below `w:bottom` — collapses both toward the body by exactly `margin − distance`
(25 px on a Word-default page) and leaves the top and bottom of the sheet blank. That was the
shipped behavior until issue #377.

## Implementation Details

### C# Changes

#### 1. New Data Attributes

Add header/footer height to section data attributes:

```csharp
// In CreateSectionDivs(), when pagination enabled
div.Add(new XAttribute("data-header-height", dims.HeaderPt.ToString("F1", NumberFormatInfo.InvariantInfo)));
div.Add(new XAttribute("data-footer-height", dims.FooterPt.ToString("F1", NumberFormatInfo.InvariantInfo)));
```

#### 2. Header/Footer Registry

When both pagination AND headers/footers are enabled, render into registry instead of document-level:

```csharp
private static XElement RenderPaginatedHeaderFooterRegistry(
    WordprocessingDocument wordDoc,
    WmlToHtmlConverterSettings settings)
{
    var registry = new XElement(Xhtml.div,
        new XAttribute("id", "pagination-hf-registry"),
        new XAttribute("style", "display:none"));

    // Get all sections and their headers/footers
    var mainDoc = wordDoc.MainDocumentPart.GetXDocument();
    var sections = GetSectionProperties(mainDoc);

    for (int sectionIndex = 0; sectionIndex < sections.Count; sectionIndex++)
    {
        var sectPr = sections[sectionIndex];

        // Render default header
        var defaultHeader = GetHeaderForType(wordDoc, sectPr, "default");
        if (defaultHeader != null)
        {
            registry.Add(WrapHeaderFooter(defaultHeader, sectionIndex, "header-default", settings));
        }

        // Render first page header (if different first page enabled)
        if (HasTitlePage(sectPr))
        {
            var firstHeader = GetHeaderForType(wordDoc, sectPr, "first");
            if (firstHeader != null)
            {
                registry.Add(WrapHeaderFooter(firstHeader, sectionIndex, "header-first", settings));
            }
        }

        // Similar for footers...
    }

    return registry;
}

private static XElement WrapHeaderFooter(
    XElement content,
    int sectionIndex,
    string hfType,
    WmlToHtmlConverterSettings settings)
{
    return new XElement(Xhtml.div,
        new XAttribute("data-section", sectionIndex),
        new XAttribute("data-hf-type", hfType),
        content);
}
```

#### 3. Conditional Rendering

Modify body transformation to use registry when both features enabled:

```csharp
if (element.Name == W.body)
{
    var bodyContent = new List<object>();

    bool paginatedHeadersFooters =
        settings.RenderHeadersAndFooters &&
        settings.RenderPagination == PaginationMode.Paginated;

    // For non-paginated mode, render headers at document level (existing behavior)
    if (settings.RenderHeadersAndFooters && !paginatedHeadersFooters)
    {
        var headersSection = RenderHeadersSection(wordDoc, settings);
        if (headersSection != null)
            bodyContent.Add(headersSection);
    }

    // ... main content ...

    // For non-paginated mode, render footers at document level (existing behavior)
    if (settings.RenderHeadersAndFooters && !paginatedHeadersFooters)
    {
        var footersSection = RenderFootersSection(wordDoc, settings);
        if (footersSection != null)
            bodyContent.Add(footersSection);
    }

    return new XElement(Xhtml.body, bodyContent);
}
```

### TypeScript Changes

#### 1. Extended Interfaces

```typescript
interface PageDimensions {
  // ... existing properties ...
  /** Header distance from top of page in points */
  headerHeight: number;
  /** Footer distance from bottom of page in points */
  footerHeight: number;
}

interface HeaderFooterEntry {
  element: HTMLElement;
  type: 'header-default' | 'header-first' | 'header-even'
      | 'footer-default' | 'footer-first' | 'footer-even';
}

interface SectionHeaderFooter {
  headerDefault?: HTMLElement;
  headerFirst?: HTMLElement;
  headerEven?: HTMLElement;
  footerDefault?: HTMLElement;
  footerFirst?: HTMLElement;
  footerEven?: HTMLElement;
}

type HeaderFooterRegistry = Map<number, SectionHeaderFooter>;
```

#### 2. Registry Parsing

```typescript
private parseHeaderFooterRegistry(): HeaderFooterRegistry {
  const registry = new Map<number, SectionHeaderFooter>();
  const registryEl = this.stagingElement.querySelector('#pagination-hf-registry');

  if (!registryEl) return registry;

  const entries = registryEl.querySelectorAll<HTMLElement>('[data-section][data-hf-type]');

  for (const entry of entries) {
    const sectionIndex = parseInt(entry.dataset.section || '0', 10);
    const hfType = entry.dataset.hfType as string;

    if (!registry.has(sectionIndex)) {
      registry.set(sectionIndex, {});
    }

    const section = registry.get(sectionIndex)!;
    const content = entry.firstElementChild as HTMLElement;

    switch (hfType) {
      case 'header-default':
        section.headerDefault = content;
        break;
      case 'header-first':
        section.headerFirst = content;
        break;
      case 'header-even':
        section.headerEven = content;
        break;
      case 'footer-default':
        section.footerDefault = content;
        break;
      case 'footer-first':
        section.footerFirst = content;
        break;
      case 'footer-even':
        section.footerEven = content;
        break;
    }
  }

  return registry;
}
```

#### 3. Header/Footer Selection Logic

```typescript
private selectHeader(
  sectionHf: SectionHeaderFooter | undefined,
  pageInSection: number,
  globalPageNumber: number
): HTMLElement | undefined {
  if (!sectionHf) return undefined;

  // First page of section uses first header if available
  if (pageInSection === 1 && sectionHf.headerFirst) {
    return sectionHf.headerFirst;
  }

  // Even pages use even header if available
  if (globalPageNumber % 2 === 0 && sectionHf.headerEven) {
    return sectionHf.headerEven;
  }

  // Default (odd) pages
  return sectionHf.headerDefault;
}

// Similar for selectFooter()
```

#### 4. Page Creation with Headers/Footers

```typescript
private createPage(
  dims: PageDimensions,
  pageNumber: number,
  sectionIndex: number,
  content: HTMLElement[],
  pageInSection: number,
  hfRegistry: HeaderFooterRegistry
): PageInfo {
  // ... existing page box and content area creation ...

  const bands = this.getPageBands(dims, sectionIndex, pageInSection, pageNumber);

  // Add header — anchored to w:header, growing downward
  const headerSource = this.selectHeader(sectionIndex, pageInSection, pageNumber);
  if (headerSource) {
    const headerDiv = document.createElement('div');
    headerDiv.className = `${this.cssPrefix}header`;
    headerDiv.style.top = `${bands.headerTop}pt`;
    headerDiv.style.height = `${bands.headerHeight}pt`;
    headerDiv.style.left = `${dims.marginLeft}pt`;
    headerDiv.style.width = `${dims.contentWidth}pt`;
    headerDiv.appendChild(headerSource.cloneNode(true) as HTMLElement);
    pageBox.appendChild(headerDiv);
  }

  // Body — between whatever the two stories left
  contentArea.style.top = `${bands.bodyTop}pt`;
  contentArea.style.height = `${bands.bodyHeight}pt`;

  // Add footer — anchored to w:footer, growing upward
  const footerSource = this.selectFooter(sectionIndex, pageInSection, pageNumber);
  if (footerSource) {
    const footerDiv = document.createElement('div');
    footerDiv.className = `${this.cssPrefix}footer`;
    footerDiv.style.bottom = `${dims.footerDistance}pt`;
    footerDiv.style.height = `${bands.footerHeight}pt`;
    footerDiv.style.left = `${dims.marginLeft}pt`;
    footerDiv.style.width = `${dims.contentWidth}pt`;
    footerDiv.appendChild(footerSource.cloneNode(true) as HTMLElement);
    pageBox.appendChild(footerDiv);
  }

  // ... rest of page creation ...
}
```

### CSS Additions

The stylesheet owns only the band's INVARIANTS; the paginator sets `top`/`bottom`/`left`/`width`/
`height` per page from the section's own distances. A stylesheet edge (`top: 0`, `bottom: 0`)
would silently win for any band the paginator left unset, so there is none.

```css
/* Paginated Header/Footer CSS */
/* w:header is the distance to the TOP of the story, which grows downward. */
.page-header {
  position: absolute;
  overflow: hidden;
  box-sizing: border-box;
  display: flex;
  flex-direction: column;
  justify-content: flex-start;
}

/* w:footer is the distance to the BOTTOM of the story, which grows upward. */
.page-footer {
  position: absolute;
  overflow: hidden;
  box-sizing: border-box;
  display: flex;
  flex-direction: column;
  justify-content: flex-end;
}

/* Hide system page number when document has its own footer */
.page-box:has(.page-footer) .page-number {
  display: none;
}

/* Print styles */
@media print {
  .page-header,
  .page-footer {
    position: static;
  }
}
```

## Configuration

### When Both Features Are Enabled

| Setting | Behavior |
|---------|----------|
| `RenderPagination = Paginated` + `RenderHeadersAndFooters = true` | Registry-based per-page headers/footers |
| `RenderPagination = None` + `RenderHeadersAndFooters = true` | Document-level headers/footers (existing) |
| `RenderPagination = Paginated` + `RenderHeadersAndFooters = false` | No headers/footers |

### No New Settings Required

The feature activates automatically when both existing settings are enabled. No new configuration options are needed.

## Dynamic Header/Footer Heights

Word documents can have headers/footers whose content exceeds the margin area. For example, a first-page header with a multi-paragraph legal disclaimer may be much taller than the standard top margin.

### Problem

If headers are constrained to the margin height:
- Content gets clipped when header is larger than margin
- Body content positioning doesn't account for actual header size
- Page breaks are calculated incorrectly

### Solution: Pre-measured Heights

The pagination engine measures all headers/footers during initialization and uses these measurements throughout:

```typescript
interface SectionHeaderFooter {
  // Element references
  headerDefault?: HTMLElement;
  headerFirst?: HTMLElement;
  headerEven?: HTMLElement;
  footerDefault?: HTMLElement;
  footerFirst?: HTMLElement;
  footerEven?: HTMLElement;

  // Pre-measured heights (populated during registry parsing)
  headerDefaultHeight?: number;
  headerFirstHeight?: number;
  headerEvenHeight?: number;
  footerDefaultHeight?: number;
  footerFirstHeight?: number;
  footerEvenHeight?: number;
}
```

### Band Resolution

For any page position, the three bands are resolved deterministically:

```typescript
private getPageBands(
  dims: PageDimensions,
  sectionIndex: number,
  pageInSection: number,
  globalPageNumber: number
): PageBands   // headerTop/headerHeight, bodyTop/bodyHeight, footerTop/footerHeight
```

It selects the story heights the page's position calls for (mirroring `selectHeader`/
`selectFooter`) and hands them to `resolvePageBands()`, which applies the margin model in
[Margin Model](#margin-model) above. This is the **single owner** of "where does anything sit
vertically on this page": band placement in `createPage`, the body budget the flow loop spends,
and the footnote area's anchor all read the same object, so they cannot disagree about where the
body ends. Consequences:

1. A story taller than its margin pushes the body instead of drawing over it
2. The note area follows the body's real bottom edge, not the raw bottom margin
3. Page breaks are calculated with the space the page actually has

A pathological story that would leave no body at all is capped, so the body band always keeps at
least `MIN_BODY_BAND_PT`; the surplus is returned to whichever band over-claimed it.

`measureHeaderFooterHeight()` therefore measures the story ALONE — no padding of its own. Padding
added there and not to the rendered band (or vice versa) is exactly how a header silently starts
overlapping body text.

### Flow Algorithm Integration

The `flowToPages()` algorithm uses dynamic content heights per page position:

```typescript
for (const block of blocks) {
  // Get the body band for the CURRENT page position
  const { bodyHeight } = getPageBands(dims, sectionIndex, pageInSection, pageNumber);

  if (blockFitsInRemainingSpace) {
    // Add to current page
  } else {
    // Start new page - recalculate for new position
    const next = getPageBands(dims, sectionIndex, pageInSection + 1, pageNumber + 1);
    remainingHeight = next.bodyHeight;
  }
}
```

### Lazy Loading Compatibility

This approach is designed for future lazy loading:

1. **Deterministic**: Same inputs always produce same page breaks
2. **Forward-only**: No backtracking during content flow
3. **Pre-computed**: Header/footer heights measured once at start
4. **Position-based**: Can compute available height for any page without knowing content

```
┌─────────────────────────────────────────────────────────────┐
│                    PAGINATION PHASES                         │
├─────────────────────────────────────────────────────────────┤
│  1. INITIALIZATION (one-time)                                │
│     └─ Parse and measure all headers/footers                │
│                                                              │
│  2. CONTENT FLOW (streamable)                                │
│     └─ Use pre-computed heights for page break decisions    │
│                                                              │
│  3. PAGE RENDERING (on-demand)                               │
│     └─ Look up cached heights, no re-measurement needed     │
└─────────────────────────────────────────────────────────────┘
```

## Limitations

1. **No section break tracking within pagination**: When content flows across section boundaries, the pagination engine doesn't detect which section a particular page belongs to (future enhancement)

2. **No dynamic field resolution**: Fields like `PAGE`, `NUMPAGES`, `DATE` in headers/footers show placeholder text, not actual values (would require JavaScript field processing)

3. **Even page headers require knowing total pages**: To correctly determine odd/even, page numbering must be sequential (currently supported)

## Future Enhancements

1. **Section boundary detection** - Track when content crosses section breaks to switch headers/footers mid-document

2. **Field resolution** - Process `PAGE`, `NUMPAGES`, `SECTIONPAGES` fields with actual page numbers

3. **Different first page per-section** - Currently supported in data model, needs pagination tracking

## Testing Strategy

1. **Unit tests**: Verify registry generation with various header/footer configurations

2. **Integration tests**: Verify pages render with correct headers/footers

3. **Visual tests**: Compare rendered output to expected layout

4. **Edge cases**:
   - Documents with no headers/footers
   - Documents with only headers (no footers)
   - Multi-section documents with different headers per section
   - First page different headers
