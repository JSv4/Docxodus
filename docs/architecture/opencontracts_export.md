# OpenContracts Export Architecture

This document describes the architecture of the OpenContracts export feature (Issue #56), which enables exporting DOCX documents to the OpenContracts format for interoperability with the OpenContracts ecosystem.

## Overview

The `OpenContractExporter` class exports DOCX documents to the `OpenContractDocExport` format, which includes:
- Complete document text content
- PAWLS-format layout information for annotation targeting
- Structural annotations (sections, paragraphs, tables)
- Relationships between structural elements

## API

### C# API

```csharp
// Export from WmlDocument
var export = OpenContractExporter.Export(wmlDoc);

// Export from WordprocessingDocument
using var wordDoc = WordprocessingDocument.Open(stream, false);
var export = OpenContractExporter.Export(wordDoc);
```

### TypeScript API

```typescript
import { exportToOpenContract } from 'docxodus';

const result = await exportToOpenContract(docxFile);
console.log(`Content: ${result.content.length} characters`);
console.log(`Pages: ${result.pageCount}`);
```

## Export Format

### OpenContractDocExport

| Field | Type | Description |
|-------|------|-------------|
| `title` | string | Document title from core properties |
| `content` | string | Complete document text (ALL content) |
| `description` | string? | Document description/subject |
| `pageCount` | int | Estimated page count |
| `pawlsFileContent` | PawlsPage[] | PAWLS layout data per page |
| `docLabels` | string[] | Document-level labels |
| `labelledText` | OpenContractsAnnotation[] | Text annotations |
| `relationships` | OpenContractsRelationship[]? | Annotation relationships |

### PAWLS Format

PAWLS (Page-Aware Layout Segmentation) provides position data for annotation targeting:

```typescript
interface PawlsPage {
  page: PawlsPageBoundary;  // Page dimensions
  tokens: PawlsToken[];      // Token positions
}

interface PawlsPageBoundary {
  width: number;   // Points (1pt = 1/72 inch)
  height: number;
  index: number;   // 0-based page index
}

interface PawlsToken {
  x: number;       // Left edge in points
  y: number;       // Top edge in points
  width: number;
  height: number;
  text: string;    // Token text (typically a word)
}
```

### Annotation Format

Annotations use text spans for positioning:

```typescript
interface OpenContractsAnnotation {
  id?: string;
  annotationLabel: string;  // "SECTION", "PARAGRAPH", "TABLE"
  rawText: string;          // Annotated text content
  page: number;             // Starting page (0-indexed)
  annotationJson?: TextSpan | Record<string, SinglePageAnnotation>;
  parentId?: string;        // Parent annotation ID
  annotationType?: string;  // "text", "structural"
  structural: boolean;      // Is structural element?
}

interface TextSpan {
  start: number;  // Start character offset
  end: number;    // End character offset (exclusive)
  text: string;
}
```

## Text Extraction

The exporter ensures 100% text coverage by extracting from all document parts:

### Main Body Content
- Paragraphs (w:p)
- Tables and nested tables (w:tbl)
- Content controls (w:sdt)
- Hyperlinks (w:hyperlink)

### Document Parts
- **Headers** (HeaderPart) - Default, first page, even page headers
- **Footers** (FooterPart) - Default, first page, even page footers
- **Footnotes** (FootnotesPart) - Excluding separators (id=0, -1)
- **Endnotes** (EndnotesPart) - Excluding separators

### Run Content
- Text (w:t)
- Tabs (w:tab) → `\t`
- Line breaks (w:br) → `\n`
- Page breaks (w:br type="page") → `\f`
- Symbols (w:sym)
- Images → `[IMAGE]` placeholder
- Objects → `[OBJECT]` placeholder

## Page Estimation

Since DOCX doesn't store explicit page boundaries, the exporter estimates pages using:

1. **Section Properties**: Page size and margins from w:sectPr
2. **Content Volume**: Characters per page heuristic (~2500 chars/page)
3. **Area Ratio**: Adjusted based on content area vs. default page size

```csharp
var areaRatio = (contentWidth * contentHeight) / defaultArea;
var adjustedCharsPerPage = (int)(avgCharsPerPage * areaRatio);
```

## Structural Annotations

The exporter automatically generates structural annotations:

| Label | Description | Parent |
|-------|-------------|--------|
| SECTION | Document section | - |
| PARAGRAPH | Text paragraph | SECTION |
| TABLE | Table element | SECTION |

### Relationships

Parent-child relationships are expressed using the `CONTAINS` relationship label:
- Section → Paragraph
- Section → Table

## Implementation Details

### File: `Docxodus/OpenContractExporter.cs`

Key methods:
- `Export(WmlDocument)` - Main entry point
- `ExtractContent()` - Extracts all text content
- `ExtractParagraph()` - Paragraph text extraction
- `ExtractTable()` - Table content extraction (recursive for nested)
- `ExtractFootnotes()` / `ExtractEndnotes()` - Note extraction
- `ExtractHeadersAndFooters()` - Header/footer extraction
- `CalculatePages()` - Page estimation algorithm
- `GeneratePawlsContent()` - PAWLS layout generation
- `GenerateStructuralAnnotations()` - Auto-generate structure annotations

### WASM Export: `wasm/DocxodusWasm/DocumentConverter.cs`

```csharp
[JSExport]
public static string ExportToOpenContract(byte[] docxBytes)
```

Returns JSON-serialized `OpenContractExportResponse`.

### TypeScript: `npm/src/index.ts`

```typescript
export async function exportToOpenContract(
  document: File | Uint8Array
): Promise<OpenContractDocExport>
```

## Usage Examples

### Complete Export

```typescript
import { exportToOpenContract } from 'docxodus';

const export = await exportToOpenContract(docxFile);

// Access complete text
console.log(`Total characters: ${export.content.length}`);

// Access structural elements
const sections = export.labelledText.filter(a => a.annotationLabel === 'SECTION');
const paragraphs = export.labelledText.filter(a => a.annotationLabel === 'PARAGRAPH');

// Access layout data
for (const page of export.pawlsFileContent) {
  console.log(`Page ${page.page.index}: ${page.page.width}x${page.page.height}pt`);
  console.log(`  ${page.tokens.length} tokens`);
}
```

### Find Text by Character Offset

```typescript
const para = export.labelledText.find(a =>
  a.annotationLabel === 'PARAGRAPH' &&
  a.annotationJson?.start <= offset &&
  a.annotationJson?.end > offset
);
```

## Compatibility

The export format is compatible with:
- [OpenContracts](https://github.com/Open-Source-Legal/OpenContracts) document analysis platform
- PAWLS (Page-Aware Layout Segmentation) annotation format
- Text-span based annotation systems

## Limitations

1. **Page Estimation**: Page boundaries are estimated, not exact
2. **Token Positions**: Token coordinates are estimated based on margins and font metrics
3. **Complex Layouts**: Text boxes, shapes, and complex positioning not fully supported
4. **Images**: Represented as `[IMAGE]` placeholder text

## Future Enhancements

- Support for text box content extraction
- More accurate token positioning using font metrics
- Support for explicit page break tracking
- Integration with PDF-based PAWLS layout extraction
