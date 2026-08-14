# Docxodus

**View, edit, and diff Word `.docx` files in the browser.** No server, no upload, no conversion
round-trip — a real OOXML engine compiled to WebAssembly, wrapped in TypeScript.

```bash
npm install docxodus
```

---

## Diff two documents into a redline Word will open

`compareDocuments()` returns a `.docx` carrying **native Word tracked changes** — `w:ins`, `w:del`,
and true move markup — not a text diff with highlighting. Open it in Word and accept/reject as usual.

![A redlined venture financing agreement](https://raw.githubusercontent.com/JSv4/Docxodus/main/docs/images/redline.png)

*Real output. One frame: a struck definition, an inserted definition, word-level substitutions inside
an untouched sentence, and a clause **moved** — struck in purple at the bottom, re-inserted at the top,
linked as one operation instead of reported as an unrelated delete and insert.*

`getRevisions()` gives you the same changes as structured data — typed, with move pairing and stable
`kind:scope:unid` anchors — so you can drive a review UI without parsing OOXML.

## Render a document faithfully

`convertDocxToHtml()` keeps what naive converters drop: justification, style inheritance, legal
numbering, tables, images, comments, headers and footers, and real footnotes with back-references.

![The NVCA model charter rendered to HTML](https://raw.githubusercontent.com/JSv4/Docxodus/main/docs/images/render.png)

Paginated mode flows content into real page boxes with per-page numbers and page-anchored footnotes —
a print-accurate preview in the browser.

Supply an exact layout token to materialize portable page citations from that layout:

```ts
const result = paginateHtml(html, viewer, {
  layoutToken: { documentVersion: session.getVersion(), rendererFingerprint: 'chromium-layout-v1' },
});
session.registerPageMap(result.pageMap!, 'chromium-layout-v1');
const citation = session.getPageCitation(anchorId, {
  documentVersion: session.getVersion(),
  rendererFingerprint: 'chromium-layout-v1',
});
navigateToPageCitation(viewer, citation);
```

Continuous renderers return explicit unavailability; Docxodus never estimates a citation page.
See the [PageMap contract](../docs/architecture/page_map.md).

## Edit it in place

`DocxEditor` is a framework-agnostic block editor over the live document. Edits go through a
`DocxSession`, and **only the changed block re-renders**, so structural operations cost ~90–360 ms on
a 346-block, 94-footnote filing template. `save()` returns lossless bytes.

![The in-browser DOCX editor](https://raw.githubusercontent.com/JSv4/Docxodus/main/docs/images/editor/editor-overview.png)

Two layers, so you can take as much UI as you want:

```ts
// The engine, with no chrome — bring your own UI.
import { DocxEditor } from 'docxodus';
const editor = DocxEditor.open(container, docxBytes, exports);

// Or the whole surface: tabbed ribbon, anchor rail, table picker, loading overlay.
import { mountRibbon } from 'docxodus';
const ribbon = mountRibbon(container, { exports });
ribbon.open(docxBytes, 'contract.docx');
```

The ribbon picks its layout from the width of the element you give it, so the same call serves a
full-page editor and a narrow embedded panel. [`examples/editor.html`](https://github.com/JSv4/Docxodus/blob/main/npm/examples/editor.html)
is a host for it in ~60 lines.

## Embed it with one script tag

The `embed` entry packages the viewer and editor behind one-call factories, with the WASM
runtime location auto-detected — no build step, no configuration:

```html
<div id="doc" style="height: 100dvh"></div>
<script type="module">
  import { createRibbonEditor } from 'https://cdn.jsdelivr.net/npm/docxodus@9.2.0/dist/embed.bundle.js';
  const ribbon = await createRibbonEditor('#doc', './contract.docx'); // or bytes / Blob / File
  // ...later: const editedBytes = ribbon.save();
</script>
```

That gives you the full editor UI. `createEditor('#doc', source, options?)` is the same document
without the chrome, and `createViewer('#doc', source, options?)` is the read-only counterpart.
Bundler users get the same API from `docxodus/embed`; classic-script pages can load
`dist/embed.iife.js` for a global `Docxodus`. See [Embedding via CDN](#embedding-via-cdn) below.

## Project it for an LLM

`convertWmlToMarkdown()` renders the document as markdown where **every block carries a stable id**,
so an agent can point at a clause and edit it — and `openDocxSession()` writes back to that same id.

Multi-step plans are atomic by default. Each callback may call any synchronous
session mutation; success is one version/undo unit, while a failure or throw
restores the complete package and history checkpoint:

```ts
const result = session.executeBatch([
  { tool: 'docx_edit', action: 'replace_text',
    mutation: () => session.replaceText(firstAnchor, 'Replacement text') },
  { tool: 'docx_create', action: 'set_header_text',
    mutation: () => session.setHeaderText(firstAnchor, 'default', 'Confidential') },
]);

if (!result.success) console.error(result.failure);
```

Pass `'best_effort'` explicitly only when partial successes should be retained.

![Markdown projection beside the rendered document](https://raw.githubusercontent.com/JSv4/Docxodus/main/docs/images/projection.png)

---

## Everything else

- **Move detection** — relocated content is identified as a move, not a delete plus an insert
- **Format change detection** — formatting-only changes (bold, italic, font size, …)
- **Comment rendering** — endnote-style, inline, or margin
- **Document metadata** — fast extraction for lazy loading and pagination
- **OpenContracts export** — for NLP and document-analysis pipelines
- **External annotations** — overlay annotations without modifying the DOCX
- **Web Worker support** — non-blocking WASM execution off the main thread
- **React hooks** — `useDocxodus`, `useConversion`, `useComparison`
- **TypeScript** — full type definitions included

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

#### First-call warmup

`createWorkerDocxodus()` warms the .NET WASM runtime, but the **comparison code
path is not exercised until your first `compareDocuments()`**. That first call
pays a one-time warmup cost (comparison-assembly initialization + JIT of the
diff/XML engine) — roughly **2× the latency** of every subsequent compare.

`prepare()` is an **optional** method that pays this cost up front. Call it once
after creating the worker — during app boot, or while the user is still picking
files — so the first user-triggered comparison is already hot. It does **not**
run automatically; if you skip it, the first compare simply absorbs the warmup
as before.

```typescript
const docxodus = await createWorkerDocxodus({ wasmBasePath: '/wasm/' });

// Optional: warm the comparison path ahead of the first user action.
await docxodus.prepare();

// Now hot — the first real compare runs at steady-state speed and triggers
// no further .wasm fetches.
const redlined = await docxodus.compareDocuments(original, modified);
```

`prepare()` is idempotent (repeated calls share one in-flight warmup and resolve
immediately once complete), needs no input documents or seed files of your own
(it builds tiny seed documents inside the worker), and is concurrent-safe —
issuing a `compareDocuments()` while a `prepare()` is still in flight will not
double-load assemblies.

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

## Embedding via CDN

Everything in the published package — the JS wrappers AND the WASM runtime — is served by npm CDNs
(jsDelivr, unpkg) with `Access-Control-Allow-Origin: *` and correct `application/wasm` MIME types,
so a page can embed a full viewer or editor without hosting anything itself.

### One-tag viewer / editor (recommended)

```html
<div id="viewer"></div>
<div id="editor"></div>
<div id="app" style="height: 100dvh"></div>
<script type="module">
  import { createViewer, createEditor, createRibbonEditor }
    from 'https://cdn.jsdelivr.net/npm/docxodus@9.2.0/dist/embed.bundle.js';

  // Choose a factory, or render several into separate containers as shown.
  // Source may be a URL, Uint8Array, ArrayBuffer, Blob, or File.
  const viewer = await createViewer('#viewer', './contract.docx');

  // No source opens a blank "New document"; returns a DocxEditor.
  const editor = await createEditor('#editor', './contract.docx', { paginated: false });
  const bytes = editor.save(); // lossless DOCX bytes

  // The whole editor UI in one call — ribbon, rail, table picker, loading overlay.
  // Density follows the CONTAINER's width, so this is also the mobile answer.
  const ribbon = await createRibbonEditor('#app', './contract.docx', {
    chrome: 'auto',       // 'full' | 'compact' | 'auto' (default)
    rail: true,           // the live kind:scope:unid / session / last-op readout
    hint: true,           // the editing hint above the document
    loader: true,         // the staged loading overlay; false removes it
  });
</script>
```

`dist/embed.bundle.js` is a self-contained ESM bundle re-exporting the entire main API
(convert, compare, diff, sessions, annotations) plus
`createViewer`/`createEditor`/`createRibbonEditor`/`mountRibbon`/`DocxEditor`, so it also works as
a no-build way to use any other function.

For pages that can't use modules, `dist/embed.iife.js` exposes the same surface as a global:

```html
<div id="doc"></div>
<script src="https://cdn.jsdelivr.net/npm/docxodus@9.2.0/dist/embed.iife.js"></script>
<script>
  Docxodus.createRibbonEditor('#doc').then((ribbon) => { /* ... */ });
</script>
```

### How WASM resolution works

The ~18 MB runtime request set in `dist/wasm/` loads lazily from the same directory the bundle
was loaded from (`import.meta.url` for module scripts, `document.currentScript` for the IIFE):
first `<bundle dir>/wasm/`, then `<bundle dir>/` as a fallback. On a CDN that resolves to
`.../dist/wasm/_framework/...` automatically. To serve the WASM assets from somewhere else, pass
`{ wasmBasePath: 'https://your.host/wasm/' }` to either factory (or call `initialize(path)` first).

### CDN caveats

- **Pin an exact version in production** (`docxodus@9.2.0`, not `@latest`) — CDN responses are
  cached as immutable, and the wire shapes between the JS wrappers and the WASM assemblies must
  come from the same release.
- The build patches the .NET loader to fetch with `credentials: "omit"` — required because the
  CDN's `Access-Control-Allow-Origin: *` cannot be combined with credentialed requests.
- The Web Worker entry (`docxodus/worker`) is not CDN-loadable cross-origin (browsers require
  same-origin worker scripts); embed runs the engine on the main thread. Self-host the package
  if you need the worker.
- Converter CSS is scoped to a private inner mount so `body`, `span`, and document-class rules
  cannot restyle the host page. Document-global `@import` and `@page` rules are omitted; use the
  full-document conversion/print path rather than an embed factory if those rules are required.
- First load fetches ~18 MB uncompressed across about 49 requests; subsequent loads hit the
  browser cache.

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
```

## Package and Runtime Size

| Component | Approx. uncompressed size |
|-----------|---------------------------:|
| dotnet.native.wasm | 1.5 MB |
| Managed/runtime WASM modules | 15.5 MB |
| WASM payload total | 17.0 MB |
| Published package (all bundles, types, maps, and WASM) | 20.0 MB |

Only the runtime request set (~18 MB including loader JavaScript) is fetched during first
initialization; source maps, declarations, and unrelated entry bundles are not. Runtime files are
loaded on demand and cached by the browser.

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
