# Docxodus

DOCX document comparison and HTML conversion in the browser using WebAssembly.

Docxodus brings professional-grade document comparison (redlining) to JavaScript applications. Compare two Word documents and get tracked changes, or convert DOCX files to HTML - all running entirely in the browser with no server required.

## Features

- **Document Comparison**: Compare two DOCX files and generate a redlined document with tracked changes
- **HTML Conversion**: Convert DOCX documents to HTML for display in the browser
- **Revision Extraction**: Get structured data about all revisions in a compared document
- **100% Client-Side**: All processing happens in the browser using WebAssembly
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
interface ConversionOptions {
  pageTitle?: string;      // HTML document title
  cssPrefix?: string;      // CSS class prefix (default: "docx-")
  fabricateClasses?: boolean; // Generate CSS classes (default: true)
  additionalCss?: string;  // Extra CSS to include
}
```

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

#### `getRevisions(document: File | Uint8Array): Promise<Revision[]>`
Extract revision information from a compared document.

```typescript
interface Revision {
  author: string;
  date: string;
  revisionType: string; // "Insertion", "Deletion", etc.
  text: string;
}
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

#### `useConversion(wasmBasePath?: string)`
Simplified hook for DOCX to HTML conversion with state management.

#### `useComparison(wasmBasePath?: string)`
Simplified hook for document comparison with state management.

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

Built on [Docxodus](https://github.com/JSv4/Redlines), a .NET library for document manipulation based on OpenXML-PowerTools.
