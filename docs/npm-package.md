# Docxodus npm Package

The `docxodus` npm package provides client-side DOCX document comparison and HTML conversion using WebAssembly. All processing runs entirely in the browser with no server required.

## Installation

```bash
npm install docxodus
```

## Features

- **Document Comparison**: Compare two DOCX files and generate a redlined document with tracked changes
- **Move Detection**: Automatically identify when content is relocated (not just deleted/inserted)
- **HTML Conversion**: Convert DOCX documents to HTML for display in the browser
- **Comment Rendering**: Render Word document comments in three different styles
- **Revision Extraction**: Get structured data about all revisions in a compared document
- **100% Client-Side**: All processing happens in the browser using WebAssembly
- **React Hooks**: Ready-to-use hooks for React applications
- **TypeScript Support**: Full type definitions included

## Quick Start

### Basic Usage

```javascript
import { initialize, convertDocxToHtml, compareDocuments } from 'docxodus';

// Initialize the WASM runtime (call once at app startup)
await initialize();

// Convert DOCX to HTML
const html = await convertDocxToHtml(docxFile);

// Compare two documents
const redlinedDocx = await compareDocuments(originalFile, modifiedFile, {
  authorName: 'Reviewer'
});
```

### React Usage

```tsx
import { useDocxodus } from 'docxodus/react';

function DocumentViewer() {
  const { isReady, isLoading, error, convertToHtml } = useDocxodus();
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

## API Reference

### Core Functions

#### `initialize(basePath?: string): Promise<void>`

Initialize the WASM runtime. Must be called before using any other functions.

By default, WASM files are auto-detected from the module's location (works with CDN, npm, or local hosting). Pass a `basePath` to load from a custom location.

#### `convertDocxToHtml(document, options?): Promise<string>`

Convert a DOCX document to HTML.

```typescript
import { convertDocxToHtml, CommentRenderMode } from 'docxodus';

const html = await convertDocxToHtml(docxFile, {
  pageTitle: 'My Document',
  cssPrefix: 'doc-',
  fabricateClasses: true,
  additionalCss: '.custom { color: red; }',
  commentRenderMode: CommentRenderMode.EndnoteStyle,
  commentCssClassPrefix: 'comment-'
});
```

**Options:**

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `pageTitle` | `string` | `"Document"` | HTML document title |
| `cssPrefix` | `string` | `"docx-"` | CSS class prefix for generated styles |
| `fabricateClasses` | `boolean` | `true` | Generate CSS classes |
| `additionalCss` | `string` | `""` | Additional CSS to include |
| `commentRenderMode` | `CommentRenderMode` | `Disabled` | How to render comments |
| `commentCssClassPrefix` | `string` | `"comment-"` | CSS prefix for comment elements |

#### `compareDocuments(original, modified, options?): Promise<Uint8Array>`

Compare two DOCX documents and return a redlined DOCX with tracked changes.

```typescript
const redlinedDocx = await compareDocuments(originalFile, modifiedFile, {
  authorName: 'Legal Team',
  detailThreshold: 0.15,
  caseInsensitive: false
});

// Save the result
const blob = new Blob([redlinedDocx], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
const url = URL.createObjectURL(blob);
```

**Options:**

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `authorName` | `string` | `"Docxodus"` | Author name for tracked changes |
| `detailThreshold` | `number` | `0.15` | 0.0-1.0, lower = more detailed comparison |
| `caseInsensitive` | `boolean` | `false` | Case-insensitive comparison |

#### `compareDocumentsToHtml(original, modified, options?): Promise<string>`

Compare documents and return the result as HTML with tracked changes rendered visually.

```typescript
const html = await compareDocumentsToHtml(originalFile, modifiedFile, {
  authorName: 'Reviewer',
  renderTrackedChanges: true  // Show <ins>/<del> elements
});
```

#### `getRevisions(document, options?): Promise<Revision[]>`

Extract revision information from a compared document. Revisions include insertions, deletions, and **moves** (relocated content).

```typescript
import { getRevisions, isMove, isMoveSource, findMovePair, RevisionType } from 'docxodus';

const revisions = await getRevisions(comparedDocx);

// Filter by type
const insertions = revisions.filter(r => r.revisionType === RevisionType.Inserted);
const deletions = revisions.filter(r => r.revisionType === RevisionType.Deleted);
const moves = revisions.filter(isMove);

// Work with move pairs
for (const rev of revisions.filter(isMoveSource)) {
  const destination = findMovePair(rev, revisions);
  console.log(`"${rev.text}" moved to "${destination?.text}"`);
}
```

**Revision Interface:**

```typescript
interface Revision {
  author: string;           // Author who made the change
  date: string;             // ISO 8601 date string
  revisionType: RevisionType | string;  // "Inserted", "Deleted", or "Moved"
  text: string;             // The revised text content
  moveGroupId?: number;     // For moves: links source and destination
  isMoveSource?: boolean;   // For moves: true = moved FROM, false = moved TO
}
```

**Helper Functions:**

| Function | Description |
|----------|-------------|
| `isInsertion(rev)` | Returns true if revision is an insertion |
| `isDeletion(rev)` | Returns true if revision is a deletion |
| `isMove(rev)` | Returns true if revision is a move |
| `isMoveSource(rev)` | Returns true if this is the source of a move (content moved FROM here) |
| `isMoveDestination(rev)` | Returns true if this is the destination of a move (content moved TO here) |
| `findMovePair(rev, allRevisions)` | Find the matching source/destination for a move revision |

### Comment Render Modes

The `CommentRenderMode` enum controls how Word document comments are rendered in HTML:

```typescript
import { CommentRenderMode } from 'docxodus';
```

| Mode | Value | Description |
|------|-------|-------------|
| `Disabled` | -1 | Don't render comments (default) |
| `EndnoteStyle` | 0 | Comments at end of document with `[1]` style bidirectional links |
| `Inline` | 1 | Tooltips via `title` and `data-comment` attributes |
| `Margin` | 2 | Side column using CSS flexbox layout |

**EndnoteStyle Example:**
```typescript
const html = await convertDocxToHtml(docxFile, {
  commentRenderMode: CommentRenderMode.EndnoteStyle
});
// Produces: highlighted text with [1] links, comments section at bottom
```

**Inline Example:**
```typescript
const html = await convertDocxToHtml(docxFile, {
  commentRenderMode: CommentRenderMode.Inline
});
// Produces: highlighted text with title="Author: comment text" attributes
```

**Margin Example:**
```typescript
const html = await convertDocxToHtml(docxFile, {
  commentRenderMode: CommentRenderMode.Margin
});
// Produces: flexbox layout with main content on left, comments in right column
```

### Move Detection

Move detection is **enabled by default** in `getRevisions()`. When content is relocated within a document, it's marked as `Moved` instead of appearing as separate deletion and insertion.

**How it works:**
- After comparison, deletions and insertions are analyzed for text similarity
- If a deletion closely matches an insertion (≥80% word overlap by default), they're linked as a move pair
- Both revisions get the same `moveGroupId` to link them together
- `isMoveSource` indicates direction: `true` = content moved FROM here, `false` = content moved TO here

**Configuration:**

Move detection can be configured via the `options` parameter to `getRevisions()`:

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `detectMoves` | `boolean` | `true` | Enable/disable move detection |
| `moveSimilarityThreshold` | `number` | `0.8` | Jaccard similarity threshold (0.0-1.0). Higher values require more exact matches. |
| `moveMinimumWordCount` | `number` | `3` | Minimum word count for move detection. Short phrases are excluded to avoid false positives. |
| `caseInsensitive` | `boolean` | `false` | When true, similarity matching ignores case differences |

**Example (detecting near-exact moves):**
```typescript
const revisions = await getRevisions(comparedDoc, {
  detectMoves: true,
  moveSimilarityThreshold: 0.95,  // Require 95% word overlap
  moveMinimumWordCount: 5         // Only consider phrases of 5+ words
});
```

**Example (detecting loose moves):**
```typescript
const revisions = await getRevisions(comparedDoc, {
  detectMoves: true,
  moveSimilarityThreshold: 0.6,   // Accept 60% word overlap
  moveMinimumWordCount: 3,
  caseInsensitive: true           // Ignore case differences
});
```

**Example (disable move detection):**
```typescript
const revisions = await getRevisions(comparedDoc, { detectMoves: false });
```

### React Hooks

#### `useDocxodus(wasmBasePath?: string)`

Main hook providing all Docxodus functionality.

```tsx
const {
  isReady,      // boolean - WASM loaded
  isLoading,    // boolean - WASM loading
  error,        // Error | null
  convertToHtml,
  compare,
  compareToHtml,
  getRevisions
} = useDocxodus();
```

#### `useConversion(wasmBasePath?: string)`

Simplified hook for DOCX to HTML conversion with state management.

```tsx
const {
  html,           // string - converted HTML
  isConverting,   // boolean
  error,          // Error | null
  convert         // (file, options?) => Promise<void>
} = useConversion();
```

#### `useComparison(wasmBasePath?: string)`

Simplified hook for document comparison with state management.

```tsx
const {
  html,           // string - comparison HTML
  result,         // Uint8Array - redlined DOCX
  isComparing,    // boolean
  error,          // Error | null
  compare,        // (original, modified, options?) => Promise<void>
  compareToHtml,  // (original, modified, options?) => Promise<void>
  downloadResult  // (filename) => void
} = useComparison();
```

## Hosting WASM Files

The WASM files are included in the npm package under `dist/wasm/`. They need to be served from your web server.

### Auto-Detection (Recommended)

By default, the library auto-detects WASM location from the module URL. This works automatically with:
- CDN usage (jsdelivr, unpkg, etc.)
- Standard npm imports in bundlers
- Direct script imports

### Manual Configuration

If auto-detection doesn't work for your setup:

```javascript
import { initialize } from 'docxodus';

// Specify custom WASM location
await initialize('/assets/wasm/');
```

### Directory Structure

After building, copy `node_modules/docxodus/dist/wasm/` to your public directory:

```
public/
  wasm/
    _framework/
      dotnet.js
      dotnet.native.wasm
      ... (other framework files)
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

## CDN Usage

You can use Docxodus directly from a CDN without npm:

```html
<script type="module">
  import { initialize, convertDocxToHtml, CommentRenderMode } from 'https://cdn.jsdelivr.net/npm/docxodus@latest/dist/index.js';

  await initialize();

  const response = await fetch('document.docx');
  const docxBytes = new Uint8Array(await response.arrayBuffer());

  const html = await convertDocxToHtml(docxBytes, {
    commentRenderMode: CommentRenderMode.EndnoteStyle
  });

  document.getElementById('content').innerHTML = html;
</script>
```

## Related Documentation

- [Comment Rendering Architecture](architecture/comment_rendering.md) - Detailed documentation on comment rendering implementation
- [DOCX Converter Architecture](architecture/docx_converter.md) - HTML conversion internals
- [Comparison Engine](architecture/comparison_engine.md) - Document comparison algorithm details

## License

MIT
