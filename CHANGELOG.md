# Changelog

All notable changes to this project will be documented in this file.

## [Unreleased] - .NET 8 / Open XML SDK 3.x Migration

### Breaking Changes
- **Target Framework**: Changed from net45/net46/netstandard2.0 to .NET 8.0
- **Open XML SDK**: Upgraded from 2.8.1 to 3.2.0
- **Graphics Library**: Replaced System.Drawing with SkiaSharp 2.88.9

### Added
- **Frame Yielding for UI Responsiveness** (Issue #44) - WASM operations now yield to the browser before heavy work begins
  - All async functions in the npm wrapper (`convertDocxToHtml`, `compareDocuments`, `compareDocumentsToHtml`, `getRevisions`, `addAnnotation`, `addAnnotationWithTarget`, `getDocumentStructure`) automatically yield using double-`requestAnimationFrame` pattern
  - This allows React state updates (loading spinners, progress indicators) to paint before blocking WASM execution
  - Transparent to consumers - no API changes required
  - Gracefully skipped in non-browser environments (Node.js, SSR)
  - Phase 1 of 3: Future phases will add Web Worker support and lazy loading
- **Custom Annotations** - Full support for adding, removing, and rendering custom annotations on DOCX documents
  - `AnnotationManager` class for programmatic annotation CRUD operations:
    - `AddAnnotation()`: Add annotation by text search or paragraph range
    - `RemoveAnnotation()`: Remove annotation by ID
    - `GetAnnotations()`: Retrieve all annotations from a document
    - `GetAnnotation()`: Get a specific annotation by ID
    - `HasAnnotations()`: Check if document has any annotations
  - `DocumentAnnotation` class with properties:
    - `Id`: Unique annotation identifier
    - `LabelId`: Category/type identifier for grouping
    - `Label`: Human-readable label text
    - `Color`: Highlight color in hex format (e.g., "#FFEB3B")
    - `Author`: Optional author name
    - `Created`: Optional creation timestamp
    - `Metadata`: Custom key-value pairs
  - `AnnotationRange` class for specifying annotation targets:
    - `FromSearch(text, occurrence)`: Find text by search
    - `FromParagraphs(start, end)`: Span paragraph indices
  - **Document Structure API** for element-based annotation targeting:
    - `DocumentStructureAnalyzer.Analyze()`: Returns navigable tree of document elements
    - `DocumentElement` class with path-based IDs (e.g., `doc/p-0`, `doc/tbl-0/tr-1/tc-2`)
    - Supported element types: `Document`, `Paragraph`, `Run`, `Table`, `TableRow`, `TableCell`, `TableColumn`, `Hyperlink`, `Image`
    - `TableColumnInfo` for virtual column elements (columns aren't real OOXML elements)
  - `AnnotationTarget` class with flexible targeting modes:
    - `Element(elementId)`: Target by element ID from structure analysis
    - `Paragraph(index)`, `ParagraphRange(start, end)`: Target by paragraph index
    - `Run(paragraphIndex, runIndex)`: Target specific run
    - `Table(index)`, `TableRow(tableIndex, rowIndex)`: Target tables/rows
    - `TableCell(tableIndex, rowIndex, cellIndex)`: Target specific cell
    - `TableColumn(tableIndex, columnIndex)`: Metadata-only column annotation
    - `TextSearch(text, occurrence)`: Search text globally
    - `SearchInElement(elementId, text, occurrence)`: Search within specific element
  - WASM methods: `GetDocumentStructure()`, `AddAnnotationWithTarget()`
  - TypeScript helper functions: `findElementById()`, `findElementsByType()`, `getParagraphs()`, `getTables()`, `getTableColumns()`
  - TypeScript targeting factories: `targetElement()`, `targetParagraph()`, `targetTableCell()`, etc.
  - React `useDocumentStructure` hook with structure navigation helpers
  - Annotations stored as Custom XML Part in DOCX (non-destructive)
  - Bookmark-based text range marking for precise positioning
  - HTML rendering with configurable label modes:
    - `AnnotationLabelMode.Above`: Floating label above highlight
    - `AnnotationLabelMode.Inline`: Label at start of highlight
    - `AnnotationLabelMode.Tooltip`: Label shown on hover
    - `AnnotationLabelMode.None`: Highlight only, no label
  - New settings in `WmlToHtmlConverterSettings`:
    - `RenderAnnotations`: Enable/disable annotation rendering
    - `AnnotationLabelMode`: Select label display mode
    - `AnnotationCssClassPrefix`: Customize CSS class names (default: "annot-")
    - `IncludeAnnotationMetadata`: Include metadata in HTML data attributes
  - WASM/npm support:
    - `getAnnotations()`, `addAnnotation()`, `removeAnnotation()`, `hasAnnotations()` functions
    - `Annotation`, `AddAnnotationRequest`, `AddAnnotationResponse`, `RemoveAnnotationResponse` types
    - `AnnotationLabelMode` enum
    - `ConversionOptions` extended with annotation rendering options
  - React support:
    - `useAnnotations` hook for annotation state management
    - `AnnotatedDocument` component with click/hover event handling
    - `useDocxodus` hook extended with annotation methods
  - 20 .NET unit tests and 21 Playwright browser tests for full coverage (including 11 for element-based targeting)
- **Comment Rendering in HTML Converter** - Full support for rendering Word document comments in HTML output
  - `CommentRenderMode` enum with three rendering modes:
    - `EndnoteStyle` (default): Comments rendered at end of document with bidirectional anchor links
    - `Inline`: Comments rendered as tooltips with `title` and `data-comment` attributes
    - `Margin`: Comments positioned in a flexbox-based margin column alongside content, with author/date headers and back-reference links
  - New settings in `WmlToHtmlConverterSettings`:
    - `RenderComments`: Enable/disable comment rendering
    - `CommentRenderMode`: Select rendering mode
    - `CommentCssClassPrefix`: Customize CSS class names (default: "comment-")
    - `IncludeCommentMetadata`: Include author/date in HTML output
  - Comment highlighting with configurable CSS classes
  - Full comment metadata support (author, date, initials)
  - Margin mode includes print-friendly CSS media queries
  - WASM/npm support via `commentRenderMode` parameter and TypeScript `CommentRenderMode` enum
- **WebAssembly NPM Package** (`docxodus`) - Browser-based document comparison and HTML conversion
  - `wasm/DocxodusWasm/` - .NET 8 WASM project with JSExport methods
  - `npm/` - TypeScript wrapper with React hooks
  - Full document comparison (redlining) support in the browser
  - DOCX to HTML conversion
  - React hooks: `useDocxodus`, `useConversion`, `useComparison`
  - Build script: `scripts/build-wasm.sh`
- **Native Move Markup in WmlComparer** - Produces Word-native move tracking markup (`w:moveFrom`/`w:moveTo`)
  - Compared documents now contain proper OpenXML move elements, not just `w:del`/`w:ins`
  - Move pairs linked via `w:name` attribute for Word compatibility
  - Range markers (`w:moveFromRangeStart`/`w:moveFromRangeEnd`, `w:moveToRangeStart`/`w:moveToRangeEnd`) properly paired
  - Microsoft Word shows moves in "Track Changes" panel as relocated content
  - New `Moved` value in `WmlComparerRevisionType` enum
  - New properties on `WmlComparerRevision`: `MoveGroupId` (links source/destination), `IsMoveSource` (true=from, false=to)
  - New settings in `WmlComparerSettings`:
    - `DetectMoves`: Enable/disable move detection (default: true)
    - `MoveSimilarityThreshold`: Jaccard similarity threshold 0.0-1.0 (default: 0.8)
    - `MoveMinimumWordCount`: Minimum words to consider for move (default: 3)
  - Uses word-level Jaccard similarity for accurate matching
  - Respects `CaseInsensitive` setting for similarity comparison
  - Full WASM/npm support with new TypeScript helpers:
    - `RevisionType.Moved` enum value
    - `isMove()`, `isMoveSource()`, `isMoveDestination()` type guards
    - `findMovePair()` function to find linked move revisions
    - `moveGroupId` and `isMoveSource` properties on `Revision` interface
- **Format Change Detection in WmlComparer** - Detects and tracks formatting-only changes (`w:rPrChange`)
  - When text content is identical but formatting changes (bold, italic, font size, etc.), produces native Word format change markup
  - Compared documents now contain `w:rPrChange` elements that Microsoft Word recognizes in Track Changes
  - New `FormatChanged` value in `WmlComparerRevisionType` enum
  - New `FormatChange` property on `WmlComparerRevision` with:
    - `OldProperties`: Dictionary of original formatting properties
    - `NewProperties`: Dictionary of new formatting properties
    - `ChangedPropertyNames`: List of what changed (e.g., "bold", "italic", "fontSize")
  - New setting in `WmlComparerSettings`:
    - `DetectFormatChanges`: Enable/disable format change detection (default: true)
  - Full WASM/npm support with new TypeScript helpers:
    - `RevisionType.FormatChanged` enum value
    - `isFormatChange()` type guard
    - `FormatChangeDetails` interface with `oldProperties`, `newProperties`, `changedPropertyNames`
    - `formatChange` property on `Revision` interface
- **Improved Revision API** - Better TypeScript support for the `getRevisions()` API
  - `RevisionType` enum with `Inserted`, `Deleted`, and `Moved` values for type-safe comparisons
  - `isInsertion()`, `isDeletion()`, `isMove()`, `isMoveSource()`, `isMoveDestination()` helper functions
  - `findMovePair()` function to find the matching revision for a move
  - Comprehensive JSDoc documentation on the `Revision` interface
  - All types are properly exported from the package
- **Paginated Headers and Footers** - Headers/footers now render correctly with pagination enabled
  - When both `RenderHeadersAndFooters` and `RenderPagination=Paginated` are enabled, headers and footers appear on each page
  - Per-section header/footer support with section index tracking
  - First page headers/footers supported (when `w:titlePg` is set in document)
  - Even page headers/footers supported for different odd/even page layouts
  - Headers/footers rendered into hidden registry for client-side cloning per-page
  - New data attributes: `data-header-height`, `data-footer-height` on section elements
  - TypeScript `PageDimensions` interface extended with `headerHeight` and `footerHeight`
  - CSS classes `.page-header` and `.page-footer` for positioning within page boxes
  - Automatic hiding of system page number when document has footer content
  - See `docs/architecture/paginated_headers_footers.md` for full architecture details
- **Per-page Footnote Rendering** - Footnotes now appear at the bottom of each page where they are referenced
  - When `RenderFootnotesAndEndnotes=true` with `RenderPagination=Paginated`, footnotes are distributed per-page
  - Footnote registry stores footnotes in a hidden container for client-side distribution
  - `data-footnote-id` attributes added to footnote references for tracking
  - Single-pass, forward-only pagination algorithm (lazy-loading compatible)
  - Pagination engine measures footnote space and includes it in page layout calculations
  - Footnotes render with separator line (`<hr>`) above them
  - **Footnote continuation**: Long footnotes that don't fit on a page are split at paragraph boundaries and continue on subsequent pages (matching Word/Office behavior)
  - **Dynamic footnote area expansion**: Footnote area can expand upward into body content space (up to 60% of page height) to fit more footnote content before splitting, reducing wasted space
  - Endnotes remain at document end (not per-page) - traditional behavior preserved
  - New TypeScript methods: `parseFootnoteRegistry()`, `extractFootnoteRefs()`, `measureFootnotesHeight()`, `addPageFootnotes()`, `splitFootnoteToFit()`, `measureContinuationHeight()`
  - New TypeScript interfaces: `FootnoteContinuation`, `PartialFootnote`
  - New TypeScript constants: `MAX_FOOTNOTE_AREA_RATIO` (0.6), `MIN_BODY_CONTENT_HEIGHT` (72pt)
  - New CSS classes: `.page-footnotes`, `.footnote-item`, `.footnote-number`, `.footnote-content`, `.footnote-continuation`
- `SkiaSharpHelpers.cs` - Color utilities for SkiaSharp compatibility
- `GetPackage()` extension method in `PtOpenXmlUtil.cs` for SDK 3.x Package access
- `SkiaSharp.NativeAssets.Linux.NoDependencies` package for Linux runtime support

### Fixed
- **Header/footer positioning in paginated mode** - Fixed headers and footers overlapping with body content. Headers now properly constrain to the top margin area (`height: marginTop`) and footers constrain to the bottom margin area (`height: marginBottom`). Uses flexbox layout for proper content alignment within constrained areas.

- **DocumentBuilder relationship copying** - Fixed bug where relationship IDs from source documents could incorrectly match existing IDs in target header/footer parts when using InsertId functionality. This caused validation errors like "The relationship 'rIdX' referenced by attribute 'r:embed' does not exist."
  - Removed flawed early-return optimization in `CopyRelatedImage()` that skipped processing when target part had matching relationship ID
  - Fixed diagram relationship handling (`R.dm`, `R.lo`, `R.qs`, `R.cs` attributes) to properly copy parts from source documents
  - Fixed chart and user shape relationship handling
  - Fixed OLE object relationship handling
  - Fixed external relationship attribute update to use correct attribute name parameter

- **SpreadsheetWriter date handling** - Fixed date cells being written with invalid ISO 8601 string format. Dates are now properly converted to Excel serial date numbers (days since December 30, 1899) which is required for transitional OOXML format.

- **WmlComparer null Unid handling** - Fixed null reference exceptions when comparing documents with elements lacking Unid attributes.

- **WmlComparer footnote/endnote comparison** (6 tests: WC-1660, WC-1670, WC-1710, WC-1720, WC-1750, WC-1760) - Fixed `AssignUnidToAllElements` to assign Unid to footnote/endnote elements themselves, enabling proper reconstruction of multi-paragraph footnotes/endnotes by `CoalesceRecurse`.

- **WmlComparer table row comparison** (1 test: WC-1500) - Added LCS-based row matching (`ApplyLcsToTableRows`) for large tables (7+ rows) when content differs significantly, preventing cascading false differences from insertions/deletions in the middle of tables.

- **WASM CDN loading CORS issue** - Fixed cross-origin loading failures when WASM files are served from CDNs (jsDelivr, unpkg). The .NET WASM runtime uses `credentials:"same-origin"` for fetch requests, which conflicts with CDN's `Access-Control-Allow-Origin: *` wildcard header. Build script now patches `dotnet.js` to use `credentials:"omit"` for CDN compatibility.

- **Vite bundler compatibility** - Added `@vite-ignore` comment to dynamic import in `npm/src/index.ts` to prevent Vite from trying to analyze/resolve the WASM loader path during development builds.

- **Pagination content overflow** - Fixed content overflowing page boundaries in the paginated view. The issue was caused by applying CSS transform scale to the content area while using inconsistent coordinate systems for positioning. The fix applies the scale transform to the entire page box instead, ensuring proper clipping and consistent scaling of all page elements.

### Changed
- Replaced `FontPartType`/`ImagePartType` with `PartTypeInfo` pattern for SDK 3.x compatibility
- Replaced `.Close()` calls with `Dispose()` pattern
- Migrated all color handling from `System.Drawing.Color` to `SKColor`
- Migrated font handling from `FontFamily`/`FontStyle` to `SKFontManager`/`SKTypeface`
- Migrated image handling from `Bitmap`/`ImageFormat` to `SKBitmap`/`SKEncodedImageFormat`

### Test Status
- 1051 passed, 0 failed, 1 skipped out of 1052 tests (~99.9% pass rate)
- Header/footer and footnote pagination changes tested via manual integration testing
