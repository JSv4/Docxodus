# Docxodus

DOCX document comparison and HTML conversion in the browser using WebAssembly.

Docxodus brings professional-grade document comparison (redlining) to JavaScript applications. Compare two Word documents and get tracked changes, or convert DOCX files to HTML - all running entirely in the browser with no server required.

## Features

- **Document Comparison**: Compare two DOCX files and generate a redlined document with tracked changes
- **Move Detection**: Automatically identifies relocated content (not just deleted/re-inserted)
- **Format Change Detection**: Detects formatting-only changes (bold, italic, font size, etc.)
- **HTML Conversion**: Convert DOCX documents to HTML for display in the browser
  - Comment rendering (endnote-style, inline, or margin)
  - Paginated output mode for PDF-like viewing
  - Headers, footers, footnotes, and endnotes support
  - Custom annotation rendering
- **Document Metadata**: Fast metadata extraction for lazy loading and pagination
- **Revision Extraction**: Get structured data about all revisions in a compared document
- **OpenContracts Export**: Export documents to OpenContracts format for NLP/document analysis
- **External Annotations**: Store annotations externally without modifying the DOCX
- **100% Client-Side**: All processing happens in the browser using WebAssembly
- **Web Worker Support**: Non-blocking WASM execution via Web Workers
- **React Hooks**: Ready-to-use hooks for React applications
- **TypeScript Support**: Full type definitions included

## Installation

```bash
npm install docxodus
```

## Quick Start

### Basic Usage

```javascript
import { initialize, convertDocxToHtml, compareDocuments } from 'docxodus';

// Initialize the WASM runtime (call once at app startup)
await initialize('/path/to/wasm/');

// Convert DOCX to HTML
const html = await convertDocxToHtml(docxFile);

// Compare two documents
const redlinedDocx = await compareDocuments(originalFile, modifiedFile, {
  authorName: 'Reviewer'
});
```

### React Usage

```tsx
import { useDocxodus, useConversion, useComparison } from 'docxodus/react';

function DocumentViewer() {
  const { isReady, isLoading, error, convertToHtml } = useDocxodus('/wasm/');
  const [html, setHtml] = useState('');

  const handleFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file && isReady) {
      const result = await convertToHtml(file);
      setHtml(result);
    }
  };

  if (isLoading) return <div>Loading...</div>;
  if (error) return <div>Error: {error.message}</div>;

  return (
    <div>
      <input type="file" accept=".docx" onChange={handleFile} />
      <div dangerouslySetInnerHTML={{ __html: html }} />
    </div>
  );
}
```

### Using the Comparison Hook

```tsx
import { useComparison } from 'docxodus/react';

function DocumentComparer() {
  const {
    html,
    isComparing,
    error,
    compareToHtml,
    downloadResult
  } = useComparison('/wasm/');

  const handleCompare = async (original: File, modified: File) => {
    await compareToHtml(original, modified, { authorName: 'Legal Team' });
  };

  return (
    <div>
      {isComparing && <p>Comparing...</p>}
      {error && <p>Error: {error.message}</p>}
      {html && <div dangerouslySetInnerHTML={{ __html: html }} />}
      <button onClick={() => downloadResult('comparison.docx')}>
        Download Redlined DOCX
      </button>
    </div>
  );
}
```

## API Reference

### Core Functions

#### `initialize(basePath?: string): Promise<void>`
Initialize the WASM runtime. Must be called before using any other functions.

#### `convertDocxToHtml(document: File | Uint8Array, options?: ConversionOptions): Promise<string>`
Convert a DOCX document to HTML.

```typescript
import { CommentRenderMode, PaginationMode, AnnotationLabelMode } from 'docxodus';

interface ConversionOptions {
  pageTitle?: string;           // HTML document title
  cssPrefix?: string;           // CSS class prefix (default: "docx-")
  fabricateClasses?: boolean;   // Generate CSS classes (default: true)
  additionalCss?: string;       // Extra CSS to include
  commentRenderMode?: CommentRenderMode;  // How to render comments (default: Disabled)
  commentCssClassPrefix?: string;         // CSS prefix for comments
  paginationMode?: PaginationMode;        // None (0) or Paginated (1)
  paginationScale?: number;               // Scale factor for pages (default: 1.0)
  renderAnnotations?: boolean;            // Render custom annotations
  annotationLabelMode?: AnnotationLabelMode;  // Above, Inline, Tooltip, or None
  renderFootnotesAndEndnotes?: boolean;   // Include footnotes/endnotes sections
  renderHeadersAndFooters?: boolean;      // Include headers and footers
  renderTrackedChanges?: boolean;         // Show insertions/deletions visually
}
```

##### Comment Render Modes

Control how Word document comments are rendered in HTML output:

```typescript
import { convertDocxToHtml, CommentRenderMode } from 'docxodus';

// Don't render comments (default)
const html = await convertDocxToHtml(docxFile, {
  commentRenderMode: CommentRenderMode.Disabled
});

// Render as footnotes with bidirectional links
const htmlEndnote = await convertDocxToHtml(docxFile, {
  commentRenderMode: CommentRenderMode.EndnoteStyle
});

// Render as inline tooltips (title attribute + data attributes)
const htmlInline = await convertDocxToHtml(docxFile, {
  commentRenderMode: CommentRenderMode.Inline
});

// Render in a side margin column (CSS flexbox layout)
const htmlMargin = await convertDocxToHtml(docxFile, {
  commentRenderMode: CommentRenderMode.Margin
});
```

| Mode | Value | Description |
|------|-------|-------------|
| `Disabled` | -1 | Don't render comments (default) |
| `EndnoteStyle` | 0 | Comments at document end with `[1]` style links |
| `Inline` | 1 | Tooltips via `title` and `data-comment` attributes |
| `Margin` | 2 | Side column using CSS flexbox |

#### `compareDocuments(original, modified, options?): Promise<Uint8Array>`
Compare two DOCX documents and return a redlined DOCX with tracked changes.

```typescript
interface CompareOptions {
  authorName?: string;     // Author name for revisions (default: "Docxodus")
  detailThreshold?: number; // 0.0-1.0, lower = more detailed (default: 0.15)
  caseInsensitive?: boolean; // Case-insensitive comparison (default: false)
}
```

#### `compareDocumentsToHtml(original, modified, options?): Promise<string>`
Compare documents and return the result as HTML.

#### `getRevisions(document: File | Uint8Array, options?): Promise<Revision[]>`
Extract revision information from a compared document.

```typescript
import {
  getRevisions,
  RevisionType,
  isInsertion,
  isDeletion,
  isMove,
  isMoveSource,
  isFormatChange,
  findMovePair
} from 'docxodus';
import type { Revision, GetRevisionsOptions } from 'docxodus';

// RevisionType enum
enum RevisionType {
  Inserted = "Inserted",      // Text or content that was added
  Deleted = "Deleted",        // Text or content that was removed
  Moved = "Moved",            // Text relocated within the document
  FormatChanged = "FormatChanged"  // Formatting-only change
}

// Revision interface with full documentation
interface Revision {
  author: string;
  date: string;
  revisionType: RevisionType | string;
  text: string;
  moveGroupId?: number;      // Links move source/destination pairs
  isMoveSource?: boolean;    // true = moved FROM here, false = moved TO here
  formatChange?: {           // Details for FormatChanged revisions
    oldProperties?: Record<string, string>;
    newProperties?: Record<string, string>;
    changedPropertyNames?: string[];
  };
}

// Get revisions with options
const revisions = await getRevisions(comparedDoc, {
  detectMoves: true,              // Enable move detection (default: true)
  moveSimilarityThreshold: 0.8,   // Jaccard similarity for moves (default: 0.8)
  moveMinimumWordCount: 3,        // Minimum words for move (default: 3)
  caseInsensitive: false          // Case-insensitive matching (default: false)
});

// Filter by type using helper functions
const insertions = revisions.filter(isInsertion);
const deletions = revisions.filter(isDeletion);
const moves = revisions.filter(isMove);
const formatChanges = revisions.filter(isFormatChange);

// Find move pairs
for (const rev of moves.filter(isMoveSource)) {
  const destination = findMovePair(rev, revisions);
  console.log(`"${rev.text}" moved to "${destination?.text}"`);
}

// Check format changes
for (const rev of formatChanges) {
  console.log(`Format changed: ${rev.formatChange?.changedPropertyNames?.join(', ')}`);
}
```

#### `getDocumentMetadata(document: File | Uint8Array): Promise<DocumentMetadata>`
Get document metadata for lazy loading and pagination without full HTML rendering.

```typescript
const metadata = await getDocumentMetadata(docxFile);

console.log(`Sections: ${metadata.sections.length}`);
console.log(`Total paragraphs: ${metadata.totalParagraphs}`);
console.log(`Estimated pages: ${metadata.estimatedPageCount}`);
console.log(`Has comments: ${metadata.hasComments}`);
console.log(`Has tracked changes: ${metadata.hasTrackedChanges}`);

// Section dimensions (in points, 1pt = 1/72 inch)
const section = metadata.sections[0];
console.log(`Page size: ${section.pageWidthPt} x ${section.pageHeightPt} pt`);
```

#### `exportToOpenContract(document: File | Uint8Array): Promise<OpenContractDocExport>`
Export document to OpenContracts format for NLP/document analysis.

```typescript
const export = await exportToOpenContract(docxFile);
console.log(`Title: ${export.title}`);
console.log(`Content: ${export.content.length} characters`);
console.log(`Pages: ${export.pageCount}`);
console.log(`Structural annotations: ${export.labelledText.length}`);
```

### Web Worker API

For non-blocking WASM execution, use the worker-based API:

```typescript
import { createWorkerDocxodus } from 'docxodus/worker';

// Create a worker instance
const docxodus = await createWorkerDocxodus({ wasmBasePath: '/wasm/' });

// All operations run in a Web Worker - main thread stays responsive
const html = await docxodus.convertDocxToHtml(docxFile, options);
const redlined = await docxodus.compareDocuments(original, modified, options);
const revisions = await docxodus.getRevisions(docxFile);
const metadata = await docxodus.getDocumentMetadata(docxFile);

// Terminate when done
docxodus.terminate();
```

### React Hooks

#### `useDocxodus(wasmBasePath?: string)`
Main hook providing all Docxodus functionality.

Returns:
- `isReady: boolean` - Whether WASM is loaded
- `isLoading: boolean` - Whether WASM is loading
- `error: Error | null` - Initialization error
- `convertToHtml()` - Convert DOCX to HTML
- `compare()` - Compare documents
- `compareToHtml()` - Compare and get HTML
- `getRevisions()` - Get revision list
- `getDocumentMetadata()` - Get document metadata

#### `useConversion(wasmBasePath?: string)`
Simplified hook for DOCX to HTML conversion with state management.

#### `useComparison(wasmBasePath?: string)`
Simplified hook for document comparison with state management.

#### `useAnnotations(wasmBasePath?: string)`
Hook for managing custom annotations on documents.

#### `useDocumentStructure(wasmBasePath?: string)`
Hook for document structure analysis and element-based targeting.

## Hosting WASM Files

The WASM files need to be served from your web server. After building:

1. Copy the contents of `dist/wasm/` to your public directory
2. Pass the path to `initialize()` or the React hooks

Example directory structure:
```
public/
  wasm/
    _framework/
      dotnet.js
      dotnet.native.wasm
      ... (other framework files)
    main.js
```

## Bundle Size

| Component | Size (uncompressed) | Size (Brotli) |
|-----------|---------------------|---------------|
| dotnet.native.wasm | ~8 MB | ~3 MB |
| Managed assemblies | ~15 MB | ~5 MB |
| Total | ~37 MB | ~10-12 MB |

The WASM files are loaded on-demand and cached by the browser.

## Browser Support

- Chrome 89+
- Firefox 89+
- Safari 15+
- Edge 89+

Requires WebAssembly SIMD support.

## License

MIT

## Credits

Built on [Docxodus](https://github.com/JSv4/Docxodus), a .NET library for document manipulation based on OpenXML-PowerTools.
