# Changelog

All notable changes to this project will be documented in this file.

## [Unreleased] - .NET 8 / Open XML SDK 3.x Migration

### Breaking Changes
- **Target Framework**: Changed from net45/net46/netstandard2.0 to .NET 8.0
- **Open XML SDK**: Upgraded from 2.8.1 to 3.2.0
- **Graphics Library**: Replaced System.Drawing with SkiaSharp 2.88.9

### Added
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
- **Move Detection in WmlComparer** - Detect relocated content as moves instead of separate deletion/insertion pairs
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
- **Improved Revision API** - Better TypeScript support for the `getRevisions()` API
  - `RevisionType` enum with `Inserted`, `Deleted`, and `Moved` values for type-safe comparisons
  - `isInsertion()`, `isDeletion()`, `isMove()`, `isMoveSource()`, `isMoveDestination()` helper functions
  - `findMovePair()` function to find the matching revision for a move
  - Comprehensive JSDoc documentation on the `Revision` interface
  - All types are properly exported from the package
- `SkiaSharpHelpers.cs` - Color utilities for SkiaSharp compatibility
- `GetPackage()` extension method in `PtOpenXmlUtil.cs` for SDK 3.x Package access
- `SkiaSharp.NativeAssets.Linux.NoDependencies` package for Linux runtime support

### Fixed
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

### Changed
- Replaced `FontPartType`/`ImagePartType` with `PartTypeInfo` pattern for SDK 3.x compatibility
- Replaced `.Close()` calls with `Dispose()` pattern
- Migrated all color handling from `System.Drawing.Color` to `SKColor`
- Migrated font handling from `FontFamily`/`FontStyle` to `SKFontManager`/`SKTypeface`
- Migrated image handling from `Bitmap`/`ImageFormat` to `SKBitmap`/`SKEncodedImageFormat`

### Test Status
- 995 passed, 0 failed, 1 skipped out of 996 tests (~99.9% pass rate)
