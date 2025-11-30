# Custom Annotations Architecture

This document describes the custom annotation system for marking and highlighting arbitrary text ranges in DOCX documents with metadata that persists in the document and renders in HTML output.

## Overview

The annotation system enables:
1. Marking arbitrary runs/paragraphs with annotations
2. **Element-based targeting** via Document Structure API (paragraphs, tables, rows, cells, runs)
3. Storing metadata: `annotation_id`, `label_id`, highlight `color` (HEX), custom key-values
4. Tracking which pages annotations span (computed at render time)
5. Rendering as highlights with floating labels in HTML
6. Persisting in DOCX without interfering with document content

## Implementation Status

| Phase | Component | Status |
|-------|-----------|--------|
| 1 | Architecture documentation | ✅ Complete |
| 2 | `AnnotationManager` C# class | ✅ Complete |
| 3 | `WmlToHtmlConverter` annotation rendering | ✅ Complete |
| 4 | .NET unit tests | ✅ Complete |
| 5 | WASM API exposure | ✅ Complete |
| 6 | TypeScript types and wrappers | ✅ Complete |
| 7 | React hooks (`useDocumentStructure`) | ✅ Complete |
| 8 | Playwright tests | ✅ Complete |
| 9 | Document Structure API | ✅ Complete |
| 10 | Element-based annotation targeting | ✅ Complete |

## Design Decisions

### Why Bookmarks + Custom XML Part?

We evaluated several DOCX extension mechanisms:

| Mechanism | Pros | Cons | Verdict |
|-----------|------|------|---------|
| Custom XML Parts | Extensible, preserved by Word | Need linking to content | ✅ Use for metadata |
| Bookmarks | Native, stable, range support | Limited metadata | ✅ Use for ranges |
| Content Controls | Rich features | Visible in Word UI | ❌ Too invasive |
| Comments | Already implemented | Shows in comment pane | ❌ Different purpose |
| Custom namespace elements | Inline | May be stripped | ❌ Risky |

**Final approach**: Bookmarks mark the text ranges, Custom XML Part stores all annotation metadata.

### Bookmark Naming Convention

```
_Docxodus_Ann_{annotation_id}
```

- Underscore prefix: Convention for system/hidden bookmarks
- `Docxodus`: Namespace to avoid collisions
- `Ann`: Short for annotation
- `{annotation_id}`: Unique identifier

Example: `_Docxodus_Ann_clause-001`

## Storage Format

### Custom XML Part

Location: `/customXml/item{N}.xml` (with corresponding `.rels` and `itemProps{N}.xml`)

```xml
<?xml version="1.0" encoding="UTF-8"?>
<annotations xmlns="http://docxodus.dev/annotations/v1" version="1.0">

  <annotation id="ann-001"
              labelId="CLAUSE_TYPE_A"
              color="#FFEB3B"
              label="Important Clause"
              created="2025-11-29T20:00:00Z"
              author="AI Model">

    <!-- Links to bookmark in document.xml -->
    <range bookmarkName="_Docxodus_Ann_ann-001"/>

    <!-- Page info (computed at render time, cached) -->
    <pageSpan startPage="1" endPage="2" stale="false"
              computedAt="2025-11-29T20:05:00Z"/>

    <!-- Extensible metadata -->
    <metadata>
      <item key="confidence">0.95</item>
      <item key="model">gpt-4</item>
      <item key="category">legal-term</item>
    </metadata>
  </annotation>

  <annotation id="ann-002"
              labelId="DATE_REF"
              color="#81C784"
              label="Date Reference">
    <range bookmarkName="_Docxodus_Ann_ann-002"/>
    <pageSpan startPage="1" endPage="1"/>
  </annotation>

</annotations>
```

### Bookmark Markers in document.xml

```xml
<!-- Single paragraph annotation -->
<w:p>
  <w:r>
    <w:t>This agreement, dated </w:t>
  </w:r>
  <w:bookmarkStart w:id="200" w:name="_Docxodus_Ann_ann-002"/>
  <w:r>
    <w:t>November 29, 2025</w:t>
  </w:r>
  <w:bookmarkEnd w:id="200"/>
  <w:r>
    <w:t>, is entered into by...</w:t>
  </w:r>
</w:p>

<!-- Multi-paragraph annotation -->
<w:p>
  <w:bookmarkStart w:id="201" w:name="_Docxodus_Ann_ann-001"/>
  <w:r>
    <w:t>WHEREAS, the Party of the First Part agrees to...</w:t>
  </w:r>
</w:p>
<w:p>
  <w:r>
    <w:t>...and furthermore commits to the following terms.</w:t>
  </w:r>
  <w:bookmarkEnd w:id="201"/>
</w:p>
```

## C# API Design

### Types

```csharp
namespace Docxodus
{
    /// <summary>
    /// Represents a custom annotation on a document range.
    /// </summary>
    public class DocumentAnnotation
    {
        /// <summary>Unique annotation identifier.</summary>
        public string Id { get; set; }

        /// <summary>Label category/type identifier (e.g., "CLAUSE_TYPE_A").</summary>
        public string LabelId { get; set; }

        /// <summary>Human-readable label text.</summary>
        public string Label { get; set; }

        /// <summary>Highlight color in hex format (e.g., "#FFEB3B").</summary>
        public string Color { get; set; }

        /// <summary>Author who created the annotation.</summary>
        public string Author { get; set; }

        /// <summary>Creation timestamp.</summary>
        public DateTime? Created { get; set; }

        /// <summary>Internal bookmark name linking to document range.</summary>
        public string BookmarkName { get; set; }

        /// <summary>Cached start page number (may be stale).</summary>
        public int? StartPage { get; set; }

        /// <summary>Cached end page number (may be stale).</summary>
        public int? EndPage { get; set; }

        /// <summary>Whether cached page info needs recalculation.</summary>
        public bool PageInfoStale { get; set; }

        /// <summary>Extensible key-value metadata.</summary>
        public Dictionary<string, string> Metadata { get; set; }

        /// <summary>The annotated text content (populated when reading).</summary>
        public string AnnotatedText { get; set; }
    }

    /// <summary>
    /// Specifies how to identify the text range for annotation.
    /// </summary>
    public class AnnotationRange
    {
        /// <summary>Search for text and annotate the Nth occurrence.</summary>
        public string SearchText { get; set; }

        /// <summary>Which occurrence to annotate (1-based). Default: 1</summary>
        public int Occurrence { get; set; } = 1;

        /// <summary>Use an existing bookmark by name.</summary>
        public string ExistingBookmarkName { get; set; }

        /// <summary>Start paragraph index (0-based).</summary>
        public int? StartParagraphIndex { get; set; }

        /// <summary>End paragraph index (0-based, inclusive).</summary>
        public int? EndParagraphIndex { get; set; }

        /// <summary>Start run index within start paragraph (0-based).</summary>
        public int? StartRunIndex { get; set; }

        /// <summary>End run index within end paragraph (0-based, inclusive).</summary>
        public int? EndRunIndex { get; set; }
    }
}
```

### AnnotationManager API

```csharp
namespace Docxodus
{
    /// <summary>
    /// Manages custom annotations in DOCX documents.
    /// </summary>
    public static class AnnotationManager
    {
        public const string AnnotationsNamespace = "http://docxodus.dev/annotations/v1";
        public const string BookmarkPrefix = "_Docxodus_Ann_";

        /// <summary>
        /// Add an annotation to a document.
        /// </summary>
        public static WmlDocument AddAnnotation(
            WmlDocument doc,
            DocumentAnnotation annotation,
            AnnotationRange range);

        /// <summary>
        /// Remove an annotation by ID.
        /// </summary>
        public static WmlDocument RemoveAnnotation(
            WmlDocument doc,
            string annotationId);

        /// <summary>
        /// Get all annotations from a document.
        /// </summary>
        public static List<DocumentAnnotation> GetAnnotations(WmlDocument doc);

        /// <summary>
        /// Get a specific annotation by ID.
        /// </summary>
        public static DocumentAnnotation GetAnnotation(WmlDocument doc, string annotationId);

        /// <summary>
        /// Update an existing annotation's metadata (not range).
        /// </summary>
        public static WmlDocument UpdateAnnotation(
            WmlDocument doc,
            DocumentAnnotation annotation);

        /// <summary>
        /// Update cached page span information for annotations.
        /// Called after pagination to store page numbers.
        /// </summary>
        public static WmlDocument UpdateAnnotationPageSpans(
            WmlDocument doc,
            Dictionary<string, (int startPage, int endPage)> pageSpans);

        /// <summary>
        /// Check if a document has any annotations.
        /// </summary>
        public static bool HasAnnotations(WmlDocument doc);

        /// <summary>
        /// Get the text content within an annotation's range.
        /// </summary>
        public static string GetAnnotatedText(WmlDocument doc, string annotationId);
    }
}
```

### WmlToHtmlConverterSettings Extensions

```csharp
public class WmlToHtmlConverterSettings
{
    // ... existing settings ...

    /// <summary>
    /// If true, render custom annotations as highlights in HTML output.
    /// Default: false
    /// </summary>
    public bool RenderAnnotations;

    /// <summary>
    /// CSS class prefix for annotation elements.
    /// Default: "annot-"
    /// </summary>
    public string AnnotationCssClassPrefix;

    /// <summary>
    /// How to display annotation labels.
    /// Default: Above
    /// </summary>
    public AnnotationLabelMode AnnotationLabelMode;

    /// <summary>
    /// If true, include annotation metadata as data attributes.
    /// Default: true
    /// </summary>
    public bool IncludeAnnotationMetadata;
}

/// <summary>
/// Specifies how annotation labels are displayed.
/// </summary>
public enum AnnotationLabelMode
{
    /// <summary>Floating label positioned above the highlight.</summary>
    Above,

    /// <summary>Label displayed inline at start of highlight.</summary>
    Inline,

    /// <summary>Label shown only on hover (tooltip).</summary>
    Tooltip,

    /// <summary>No labels displayed, only highlights.</summary>
    None
}
```

## HTML Output

### Structure

```html
<!-- Annotation with floating label -->
<span class="annot-highlight"
      data-annotation-id="ann-001"
      data-label-id="CLAUSE_TYPE_A"
      data-label="Important Clause"
      data-author="AI Model"
      data-created="2025-11-29T20:00:00Z"
      style="--annot-color: #FFEB3B;">
  <span class="annot-label">Important Clause</span>
  WHEREAS, the Party of the First Part agrees to...
</span>

<!-- Multi-paragraph annotation spans multiple elements -->
<p>
  <span class="annot-highlight annot-start"
        data-annotation-id="ann-001"
        style="--annot-color: #FFEB3B;">
    <span class="annot-label">Important Clause</span>
    First paragraph of annotated content...
  </span>
</p>
<p>
  <span class="annot-highlight annot-continuation"
        data-annotation-id="ann-001"
        style="--annot-color: #FFEB3B;">
    Continuation of annotated content...
  </span>
</p>
<p>
  <span class="annot-highlight annot-end"
        data-annotation-id="ann-001"
        style="--annot-color: #FFEB3B;">
    Final paragraph of annotated content.
  </span>
</p>
```

### Generated CSS

```css
/* Custom Annotations CSS */

/* Annotation highlight base */
.annot-highlight {
    position: relative;
    display: inline;
    background-color: color-mix(in srgb, var(--annot-color, #FFFF00) 35%, transparent);
    border-bottom: 2px solid var(--annot-color, #FFFF00);
    padding: 1px 2px;
    border-radius: 2px;
    transition: background-color 0.15s ease;
}

.annot-highlight:hover {
    background-color: color-mix(in srgb, var(--annot-color, #FFFF00) 50%, transparent);
}

/* Floating label above highlight */
.annot-label {
    position: absolute;
    top: -1.7em;
    left: 0;
    font-size: 0.7em;
    font-weight: 600;
    background: var(--annot-color, #FFFF00);
    color: #000;
    padding: 2px 6px;
    border-radius: 3px;
    white-space: nowrap;
    box-shadow: 0 1px 3px rgba(0,0,0,0.2);
    z-index: 100;
    pointer-events: none;
    line-height: 1.2;
}

/* Dark text for light backgrounds, light text for dark */
.annot-label {
    /* Could use contrast calculation in JS for better results */
}

/* Only show label on first segment of multi-paragraph annotations */
.annot-continuation .annot-label,
.annot-end .annot-label {
    display: none;
}

/* Inline label mode */
.annot-highlight[data-label-mode="inline"] .annot-label {
    position: static;
    display: inline;
    margin-right: 4px;
    font-size: 0.8em;
    vertical-align: middle;
}

/* Tooltip mode - show on hover */
.annot-highlight[data-label-mode="tooltip"] .annot-label {
    display: none;
    top: auto;
    bottom: 100%;
    margin-bottom: 4px;
}

.annot-highlight[data-label-mode="tooltip"]:hover .annot-label {
    display: block;
}

/* No label mode */
.annot-highlight[data-label-mode="none"] .annot-label {
    display: none;
}

/* Handle nested/overlapping annotations */
.annot-highlight .annot-highlight {
    background: none;
    border-bottom-style: dashed;
    padding: 0;
}

/* Ensure labels don't overlap badly */
.annot-highlight .annot-highlight .annot-label {
    top: -3.2em;
}
```

## TypeScript Types

```typescript
// npm/src/types.ts

/**
 * Represents a custom annotation on a document range.
 */
export interface DocumentAnnotation {
  /** Unique annotation identifier */
  id: string;
  /** Label category/type identifier */
  labelId: string;
  /** Human-readable label text */
  label: string;
  /** Highlight color in hex format (e.g., "#FFEB3B") */
  color: string;
  /** Author who created the annotation */
  author?: string;
  /** Creation timestamp (ISO 8601) */
  created?: string;
  /** Cached start page number */
  startPage?: number;
  /** Cached end page number */
  endPage?: number;
  /** Whether cached page info needs recalculation */
  pageInfoStale?: boolean;
  /** Extensible key-value metadata */
  metadata?: Record<string, string>;
  /** The annotated text content */
  annotatedText?: string;
}

/**
 * Specifies how to identify the text range for annotation.
 */
export interface AnnotationRange {
  /** Search for text and annotate the Nth occurrence */
  searchText?: string;
  /** Which occurrence to annotate (1-based, default: 1) */
  occurrence?: number;
  /** Use an existing bookmark by name */
  existingBookmarkName?: string;
  /** Start paragraph index (0-based) */
  startParagraphIndex?: number;
  /** End paragraph index (0-based, inclusive) */
  endParagraphIndex?: number;
  /** Start run index within start paragraph */
  startRunIndex?: number;
  /** End run index within end paragraph */
  endRunIndex?: number;
}

/**
 * How annotation labels are displayed.
 */
export enum AnnotationLabelMode {
  /** Floating label positioned above the highlight */
  Above = 0,
  /** Label displayed inline at start of highlight */
  Inline = 1,
  /** Label shown only on hover (tooltip) */
  Tooltip = 2,
  /** No labels displayed, only highlights */
  None = 3
}

/**
 * Options for annotation rendering.
 */
export interface AnnotationRenderOptions {
  /** CSS class prefix (default: "annot-") */
  cssPrefix?: string;
  /** Label display mode (default: Above) */
  labelMode?: AnnotationLabelMode;
  /** Include metadata as data attributes (default: true) */
  includeMetadata?: boolean;
}
```

## React Components

```tsx
// npm/src/react.tsx

/**
 * Props for AnnotatedDocument component.
 */
export interface AnnotatedDocumentProps {
  /** HTML content with annotation markup */
  html: string;
  /** Annotations data (for callbacks/interactivity) */
  annotations?: DocumentAnnotation[];
  /** Callback when annotation is clicked */
  onAnnotationClick?: (annotation: DocumentAnnotation) => void;
  /** Callback when annotation is hovered */
  onAnnotationHover?: (annotation: DocumentAnnotation | null) => void;
  /** Label display mode override */
  labelMode?: AnnotationLabelMode;
  /** Additional CSS class for container */
  className?: string;
  /** Inline styles for container */
  style?: React.CSSProperties;
}

/**
 * Renders HTML content with interactive annotations.
 */
export function AnnotatedDocument(props: AnnotatedDocumentProps): JSX.Element;

/**
 * Hook for managing annotations on a document.
 */
export function useAnnotations(docxBytes: Uint8Array | null): {
  annotations: DocumentAnnotation[];
  isLoading: boolean;
  error: Error | null;
  addAnnotation: (annotation: DocumentAnnotation, range: AnnotationRange) => Promise<Uint8Array>;
  removeAnnotation: (annotationId: string) => Promise<Uint8Array>;
  updateAnnotation: (annotation: DocumentAnnotation) => Promise<Uint8Array>;
};
```

## Page Span Tracking

Page numbers are dynamic in DOCX. We compute them at render time:

### During Pagination

```typescript
interface PaginationResult {
  totalPages: number;
  pages: PageInfo[];
  /** Annotation page spans computed during pagination */
  annotationSpans: AnnotationPageSpan[];
}

interface AnnotationPageSpan {
  annotationId: string;
  startPage: number;
  endPage: number;
  /** Bounding rectangles on each page (for precise label positioning) */
  pageRects: Map<number, DOMRect[]>;
}
```

### Caching Strategy

1. Compute page spans during pagination
2. Store in Custom XML Part with `stale="false"` and `computedAt` timestamp
3. On document edit (detected by checksum change), mark as `stale="true"`
4. Re-compute on next render

## Test Coverage Plan

### .NET Unit Tests (`Docxodus.Tests/AnnotationManagerTests.cs`)

1. **CRUD Operations**
   - `AddAnnotation_WithSearchText_CreatesBookmarkAndCustomXml`
   - `AddAnnotation_WithParagraphRange_CreatesCorrectBookmark`
   - `AddAnnotation_WithExistingBookmark_LinksToBookmark`
   - `RemoveAnnotation_DeletesBookmarkAndCustomXml`
   - `GetAnnotations_ReturnsAllAnnotations`
   - `GetAnnotation_ById_ReturnsCorrectAnnotation`
   - `UpdateAnnotation_UpdatesMetadataPreservesRange`
   - `HasAnnotations_ReturnsTrueWhenPresent`
   - `HasAnnotations_ReturnsFalseWhenEmpty`

2. **Edge Cases**
   - `AddAnnotation_MultiParagraphRange_SpansCorrectly`
   - `AddAnnotation_DuplicateId_ThrowsException`
   - `RemoveAnnotation_NonexistentId_NoOp`
   - `GetAnnotatedText_ReturnsCorrectContent`
   - `AddAnnotation_WithMetadata_StoresAllKeyValues`

3. **Round-Trip**
   - `Annotation_SurvivesSerializationRoundTrip`
   - `Annotation_PreservedAfterDocumentModification`

### HTML Converter Tests (`Docxodus.Tests/HtmlConverterTests.cs`)

1. **Rendering**
   - `ConvertToHtml_WithAnnotations_RendersHighlights`
   - `ConvertToHtml_WithAnnotations_IncludesDataAttributes`
   - `ConvertToHtml_WithAnnotations_GeneratesCss`
   - `ConvertToHtml_LabelModeAbove_RendersFloatingLabels`
   - `ConvertToHtml_LabelModeTooltip_RendersHoverLabels`
   - `ConvertToHtml_LabelModeNone_NoLabels`

2. **Multi-Paragraph**
   - `ConvertToHtml_MultiParagraphAnnotation_SpansElements`
   - `ConvertToHtml_OverlappingAnnotations_HandledCorrectly`

### Playwright Tests (`npm/tests/docxodus.spec.ts`)

1. **WASM API**
   - `addAnnotation creates annotation in document`
   - `getAnnotations returns all annotations`
   - `removeAnnotation deletes annotation`
   - `annotation survives document round-trip`

2. **HTML Rendering**
   - `annotations render as highlights with correct colors`
   - `annotation labels display above highlights`
   - `annotation data attributes are present`
   - `clicking annotation triggers callback`
   - `hovering annotation shows tooltip (tooltip mode)`

3. **Pagination Integration**
   - `annotation page spans are calculated correctly`
   - `multi-page annotation reports correct start/end pages`

4. **Document Structure API**
   - `can get document structure`
   - `structure contains paragraphs with correct IDs`
   - `structure contains tables and columns`
   - `structure contains rows and cells for tables`
   - `can search within structure`

5. **Element-based Annotation Targeting**
   - `can target paragraph by index`
   - `can target table cell by indices`
   - `can target by element ID`
   - `can search text within element`
   - `table column annotations store metadata`
   - `multiple targeting methods work together`

## Document Structure API

The Document Structure API enables precise targeting of document elements for annotation. It analyzes DOCX documents and returns a navigable tree of typed elements with stable IDs.

### Element Types

```csharp
public enum DocumentElementType
{
    Document,   // Root document
    Paragraph,  // w:p elements
    Run,        // w:r elements within paragraphs
    Table,      // w:tbl elements
    TableRow,   // w:tr elements
    TableCell,  // w:tc elements
    TableColumn,// Virtual - metadata only (columns aren't real OOXML elements)
    Hyperlink,  // w:hyperlink elements
    Image       // w:drawing/w:pict elements
}
```

### Element ID Format

Element IDs are path-based, stable identifiers:

```
doc                     # Document root
doc/p-0                 # First paragraph
doc/p-0/r-0             # First run in first paragraph
doc/p-0/hl-0            # First hyperlink in first paragraph
doc/tbl-0               # First table
doc/tbl-0/tr-0          # First row in first table
doc/tbl-0/tr-0/tc-0     # First cell in first row
doc/tbl-0/tr-0/tc-0/p-0 # First paragraph in that cell
```

### C# API

```csharp
namespace Docxodus
{
    /// <summary>
    /// Represents a document element in the structure tree.
    /// </summary>
    public class DocumentElement
    {
        public string Id { get; set; }
        public DocumentElementType Type { get; set; }
        public string? TextPreview { get; set; }  // First ~100 chars
        public int Index { get; set; }            // Position in parent
        public List<DocumentElement> Children { get; set; }

        // Table-specific properties
        public int? RowIndex { get; set; }
        public int? ColumnIndex { get; set; }
        public int? RowSpan { get; set; }
        public int? ColumnSpan { get; set; }
    }

    /// <summary>
    /// Table column information (virtual element).
    /// </summary>
    public class TableColumnInfo
    {
        public string TableId { get; set; }
        public int ColumnIndex { get; set; }
        public List<string> CellIds { get; set; }
        public int RowCount { get; set; }
    }

    /// <summary>
    /// Complete document structure.
    /// </summary>
    public class DocumentStructure
    {
        public DocumentElement Root { get; set; }
        public Dictionary<string, DocumentElement> ElementsById { get; set; }
        public Dictionary<string, TableColumnInfo> TableColumns { get; set; }
    }

    /// <summary>
    /// Analyzes document structure.
    /// </summary>
    public static class DocumentStructureAnalyzer
    {
        public static DocumentStructure Analyze(WmlDocument doc);
    }
}
```

### TypeScript Types

```typescript
export enum DocumentElementType {
  Document = 'Document',
  Paragraph = 'Paragraph',
  Run = 'Run',
  Table = 'Table',
  TableRow = 'TableRow',
  TableCell = 'TableCell',
  TableColumn = 'TableColumn',
  Hyperlink = 'Hyperlink',
  Image = 'Image'
}

export interface DocumentElement {
  Id: string;
  Type: string;
  TextPreview?: string;
  Index: number;
  Children: DocumentElement[];
  RowIndex?: number;
  ColumnIndex?: number;
  RowSpan?: number;
  ColumnSpan?: number;
}

export interface DocumentStructure {
  root: DocumentElement;
  elementsById: Record<string, DocumentElement>;
  tableColumns: Record<string, TableColumnInfo>;
}

// Helper functions
export function findElementById(structure: DocumentStructure, id: string): DocumentElement | undefined;
export function findElementsByType(structure: DocumentStructure, type: DocumentElementType): DocumentElement[];
export function getParagraphs(structure: DocumentStructure): DocumentElement[];
export function getTables(structure: DocumentStructure): DocumentElement[];
export function getTableColumns(structure: DocumentStructure, tableId: string): TableColumnInfo[];
```

## Element-based Annotation Targeting

The `AnnotationTarget` class provides flexible targeting modes for annotations.

### Targeting Modes

```csharp
public enum AnnotationTargetMode
{
    /// <summary>Target by element ID from Document Structure API.</summary>
    ElementId,

    /// <summary>Search for text globally in document.</summary>
    TextSearch,

    /// <summary>Target by element type and indices (e.g., paragraph 3, table cell [1,2,3]).</summary>
    IndexBased,

    /// <summary>Target a table column (metadata-only, columns aren't real elements).</summary>
    TableColumn,

    /// <summary>Search for text within a specific element.</summary>
    SearchInElement
}
```

### AnnotationTarget API

```csharp
public class AnnotationTarget
{
    // Targeting mode (auto-detected from which properties are set)
    public AnnotationTargetMode GetTargetMode();

    // Element ID targeting
    public string? ElementId { get; set; }

    // Index-based targeting
    public DocumentElementType? ElementType { get; set; }
    public int? ParagraphIndex { get; set; }
    public int? RunIndex { get; set; }
    public int? TableIndex { get; set; }
    public int? RowIndex { get; set; }
    public int? CellIndex { get; set; }
    public int? ColumnIndex { get; set; }

    // Text search targeting
    public string? SearchText { get; set; }
    public int Occurrence { get; set; } = 1;

    // Paragraph range targeting
    public int? RangeEndParagraphIndex { get; set; }

    // Factory methods
    public static AnnotationTarget Element(string elementId);
    public static AnnotationTarget Paragraph(int index);
    public static AnnotationTarget ParagraphRange(int startIndex, int endIndex);
    public static AnnotationTarget Run(int paragraphIndex, int runIndex);
    public static AnnotationTarget Table(int tableIndex);
    public static AnnotationTarget TableRow(int tableIndex, int rowIndex);
    public static AnnotationTarget TableCell(int tableIndex, int rowIndex, int cellIndex);
    public static AnnotationTarget TableColumn(int tableIndex, int columnIndex);
    public static AnnotationTarget TextSearch(string text, int occurrence = 1);
    public static AnnotationTarget SearchInElement(string elementId, string text, int occurrence = 1);
}
```

### Usage Examples

```csharp
// Target a specific paragraph
var target1 = AnnotationTarget.Paragraph(2);

// Target a table cell
var target2 = AnnotationTarget.TableCell(tableIndex: 0, rowIndex: 1, cellIndex: 2);

// Target by element ID from structure analysis
var structure = DocumentStructureAnalyzer.Analyze(doc);
var cellId = structure.ElementsById.Keys.First(k => k.Contains("/tc-"));
var target3 = AnnotationTarget.Element(cellId);

// Search within a specific element
var target4 = AnnotationTarget.SearchInElement("doc/p-0", "important", occurrence: 1);

// Add annotation with target
var annotation = new DocumentAnnotation { Id = "ann-1", Label = "My Label", Color = "#FFEB3B" };
var result = AnnotationManager.AddAnnotation(doc, annotation, target1);
```

### TypeScript Usage

```typescript
import {
  getDocumentStructure,
  addAnnotationWithTarget,
  targetParagraph,
  targetTableCell,
  targetElement,
  findElementsByType,
  DocumentElementType
} from 'docxodus';

// Get document structure
const structure = await getDocumentStructure(docxBytes);

// Find all paragraphs
const paragraphs = findElementsByType(structure, DocumentElementType.Paragraph);

// Add annotation to first paragraph
const result = await addAnnotationWithTarget(docxBytes, {
  Id: 'ann-1',
  LabelId: 'IMPORTANT',
  Label: 'Important Section',
  Color: '#FFEB3B',
  ...targetParagraph(0)
});

// Add annotation to table cell
const cellResult = await addAnnotationWithTarget(docxBytes, {
  Id: 'ann-2',
  LabelId: 'DATA',
  Label: 'Key Data',
  Color: '#81C784',
  ...targetTableCell(0, 1, 2)  // table 0, row 1, cell 2
});

// Add annotation by element ID
const result3 = await addAnnotationWithTarget(docxBytes, {
  Id: 'ann-3',
  LabelId: 'REF',
  Label: 'Reference',
  Color: '#64B5F6',
  ...targetElement(paragraphs[0].Id)
});
```

### React Hook

```typescript
import { useDocumentStructure } from 'docxodus/react';

function DocumentViewer({ docxBytes }: { docxBytes: Uint8Array }) {
  const {
    structure,
    isLoading,
    error,
    paragraphs,
    tables,
    findById,
    getTableColumns
  } = useDocumentStructure(docxBytes);

  if (isLoading) return <div>Loading...</div>;
  if (error) return <div>Error: {error.message}</div>;

  return (
    <div>
      <h3>Paragraphs ({paragraphs.length})</h3>
      <ul>
        {paragraphs.map(p => (
          <li key={p.Id}>
            {p.Id}: {p.TextPreview?.slice(0, 50)}...
          </li>
        ))}
      </ul>

      <h3>Tables ({tables.length})</h3>
      <ul>
        {tables.map(t => (
          <li key={t.Id}>
            {t.Id}: {t.Children.length} rows
          </li>
        ))}
      </ul>
    </div>
  );
}
```

## Table Column Annotations

Table columns are **virtual elements** - they don't exist as discrete OOXML elements. Column annotations are stored as metadata only:

- No bookmark is created in the document
- Annotation metadata includes `ColumnIndex` and `TableIndex`
- The annotation's `BookmarkName` is left empty
- Column-scoped highlights can be applied in HTML by annotating all cells in the column

### Column Annotation Storage

```xml
<annotation id="col-ann-1"
            labelId="COLUMN_HEADER"
            color="#E1BEE7"
            label="Price Column">
  <!-- No bookmark - columns are virtual -->
  <columnTarget tableIndex="0" columnIndex="2"/>
  <metadata>
    <item key="columnType">currency</item>
  </metadata>
</annotation>
```

## File Changes Summary

| File | Change Type | Description |
|------|-------------|-------------|
| `Docxodus/AnnotationManager.cs` | New | Core annotation CRUD + element targeting |
| `Docxodus/DocumentAnnotation.cs` | New | Annotation data types |
| `Docxodus/AnnotationTarget.cs` | New | Flexible targeting modes |
| `Docxodus/DocumentStructure.cs` | New | Document structure analyzer |
| `Docxodus/WmlToHtmlConverter.cs` | Modify | Add annotation rendering |
| `Docxodus.Tests/AnnotationManagerTests.cs` | New | Unit tests |
| `wasm/DocxodusWasm/DocumentConverter.cs` | Modify | WASM exports for annotations |
| `wasm/DocxodusWasm/JsonContext.cs` | Modify | JSON serialization types |
| `npm/src/types.ts` | Modify | TypeScript types + targeting helpers |
| `npm/src/index.ts` | Modify | Wrapper functions for structure/targeting |
| `npm/src/react.ts` | Modify | `useDocumentStructure` hook |
| `npm/dist/wasm/test-harness.html` | Modify | Test helper functions |
| `npm/tests/docxodus.spec.ts` | Modify | Playwright tests |
| `docs/architecture/custom_annotations.md` | Modify | This document |
| `CHANGELOG.md` | Modify | Document feature |
