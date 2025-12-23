# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Important Instructions

- **Never credit yourself in commits.** Do not add "Generated with Claude Code" or "Co-Authored-By: Claude" to commit messages.

## Coding Standards

### Nullable Reference Types

The project has `<Nullable>disable</Nullable>` globally due to ~9,000 warnings in legacy code. However, **new code should use nullable annotations**:

- **New files**: Add `#nullable enable` at the top of the file
- **Substantial refactors**: When significantly modifying an existing file, consider adding `#nullable enable` and fixing warnings in that file
- **Use proper annotations**: Mark nullable parameters/returns with `?`, use null checks or `!` where appropriate

```csharp
#nullable enable

namespace Docxodus;

public class MyNewClass
{
    public string Name { get; set; } = string.Empty;  // Non-nullable with default
    public string? Description { get; set; }           // Explicitly nullable

    public string? FindItem(string key)                // May return null
    {
        // ...
    }
}
```

See [Issue #13](https://github.com/JSv4/Docxodus/issues/13) for the full nullable migration plan.

## Feature Development Workflow

When implementing new features or significant changes, follow this workflow:

### 1. Documentation Updates

- **CHANGELOG.md** - Add entry under `[Unreleased]` section describing the feature/fix
- **CLAUDE.md** - Update if the feature adds new settings, modules, or changes architecture
- **docs/architecture/** - Create or update architecture docs for significant features (e.g., `comment_rendering.md`, `comparison_engine.md`)
- **docs/ooxml_corner_cases.md** - Document any OOXML edge cases where Word's behavior differs from spec or our implementation (see below)

### 2. Test Updates

- Add tests to the appropriate test file in `Docxodus.Tests/`:
  - `HtmlConverterTests.cs` - WmlToHtmlConverter features
  - `WmlComparerTests.cs` - Document comparison features
  - `DocumentBuilderTests.cs` - Document merging/splitting
  - Use existing test files from `TestFiles/` when possible
  - When creating programmatic test documents, ensure all required parts exist (StyleDefinitionsPart, DocumentSettingsPart, etc.)

### 3. WASM/npm Wrapper Updates

Update these when adding new settings or methods to the .NET API:

- **wasm/DocxodusWasm/DocumentConverter.cs** - Add new JSExport methods or parameters
- **wasm/DocxodusWasm/DocumentComparer.cs** - For comparison-related changes
- **npm/src/types.ts** - Add TypeScript types, enums, and update `DocxodusWasmExports` interface
- **npm/src/index.ts** - Update wrapper functions to use new WASM methods

Build and verify with:
```bash
npm run build          # Builds WASM and TypeScript
dotnet test            # Run .NET tests
```

### 4. When to Update Each Layer

| Change Type | .NET | Tests | WASM | npm/TS | Docs |
|-------------|------|-------|------|--------|------|
| New converter setting | ✓ | ✓ | ✓ | ✓ | ✓ |
| Bug fix | ✓ | ✓ | - | - | CHANGELOG |
| New public enum | ✓ | ✓ | ✓ | ✓ | ✓ |
| Internal refactor | ✓ | ✓ | - | - | - |
| New module | ✓ | ✓ | ✓ | ✓ | ✓ |

## Build Commands

```bash
# Build the entire solution
dotnet build Docxodus.sln

# Build specific project
dotnet build Docxodus/Docxodus.csproj
```

## Test Commands

```bash
# Run all tests
dotnet test Docxodus.Tests/Docxodus.Tests.csproj

# Run a specific test by name
dotnet test --filter "FullyQualifiedName~DB001_DocumentBuilderKeepSections"

# Run tests for a specific test class
dotnet test --filter "FullyQualifiedName~DbTests"
```

## Architecture Overview

Docxodus is a library for manipulating Open XML documents (DOCX, XLSX, PPTX) built on top of the Open XML SDK. It is a fork of OpenXmlPowerTools upgraded to .NET 8.0. All code is in the `Docxodus` namespace.

### Document Wrapper Classes

The library uses in-memory byte array wrappers for documents:
- `DocxodusDocument` - Base class holding `DocumentByteArray` and `FileName`
- `WmlDocument` - Word documents (.docx)
- `SmlDocument` - Spreadsheet documents (.xlsx)
- `PmlDocument` - Presentation documents (.pptx)

These allow immutable-style document manipulation via `OpenXmlMemoryStreamDocument` pattern:
```csharp
using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(doc))
{
    using (WordprocessingDocument document = streamDoc.GetWordprocessingDocument())
    {
        // modify document
    }
    return streamDoc.GetModifiedWmlDocument();
}
```

### Core Modules

**DocumentBuilder.cs** - Merge/split DOCX files. Uses `Source` objects to specify document ranges:
```csharp
var sources = new List<Source> { new Source(wmlDoc, keepSections: true) };
DocumentBuilder.BuildDocument(sources, outputPath);
```

**WmlComparer.cs** - Compare two DOCX files, producing a document with tracked revisions. Supports nested tables and text boxes. Key settings in `WmlComparerSettings`:
- `AuthorForRevisions` - Author name for tracked changes
- `DetailThreshold` - 0.0-1.0, lower = more detailed comparison (default: 0.15)
- `CaseInsensitive` - Case-insensitive comparison
- `DetectMoves` - Enable move detection in `GetRevisions()` (default: true)
- `MoveSimilarityThreshold` - Jaccard similarity threshold for moves (default: 0.8)
- `MoveMinimumWordCount` - Minimum words for move detection (default: 3)
- `DetectFormatChanges` - Enable format change detection (default: true)

Move detection produces **native Word move markup** (`w:moveFrom`/`w:moveTo`) when `DetectMoves` is enabled:
- The comparer analyzes deleted/inserted content blocks for similarity after LCS comparison
- Matching pairs (≥80% Jaccard similarity by default) are converted to move markup
- The output document contains `w:moveFromRangeStart`/`w:moveFromRangeEnd` and `w:moveToRangeStart`/`w:moveToRangeEnd` elements
- Move pairs are linked via the `w:name` attribute (e.g., "move1")
- `GetRevisions()` recognizes this native markup and returns `WmlComparerRevisionType.Moved` revisions
- `WmlComparerRevision.MoveGroupId` links source and destination revisions
- `WmlComparerRevision.IsMoveSource` - true = moved FROM here, false = moved TO here

Format change detection produces **native Word format change markup** (`w:rPrChange`) when `DetectFormatChanges` is enabled:
- The comparer analyzes Equal atoms (same text content) for run property differences after LCS comparison
- When text is identical but formatting differs (bold, italic, font size, etc.), atoms are marked as FormatChanged
- The output document contains `w:rPrChange` elements inside `w:rPr` with the old formatting properties
- `GetRevisions()` recognizes this native markup and returns `WmlComparerRevisionType.FormatChanged` revisions
- `WmlComparerRevision.FormatChange` contains details about what changed (old/new properties, changed property names)

**WmlToHtmlConverter.cs / HtmlToWmlConverter.cs** - Bidirectional DOCX ↔ HTML conversion. Key settings in `WmlToHtmlConverterSettings`:
- `RenderTrackedChanges` - Render insertions/deletions as `<ins>`/`<del>` instead of accepting them
- `RenderMoveOperations` - Distinguish move operations from regular insert/delete
- `RenderFootnotesAndEndnotes` - Include footnotes/endnotes sections in HTML output
- `RenderHeadersAndFooters` - Include document headers/footers in HTML output
- `RenderComments` - Render document comments in HTML output
- `CommentRenderMode` - How to render comments: `EndnoteStyle` (default), `Inline`, or `Margin`
- `AuthorColors` - Dictionary mapping author names to CSS colors for styling

See `docs/architecture/comment_rendering.md` for detailed comment rendering documentation.

**DocumentAssembler.cs** - Template population from XML data using content controls.

**PresentationBuilder.cs** - Merge/split PPTX files.

**SpreadsheetWriter.cs** - Simplified XLSX creation API with streaming support for large files.

**OpenXmlRegex.cs** - Search/replace in DOCX/PPTX using regular expressions.

**RevisionAccepter.cs / RevisionProcessor.cs** - Handle tracked revisions.

**FormattingAssembler.cs** - Resolve and flatten document formatting.

**MetricsGetter.cs** - Extract document metrics (styles, fonts, languages).

**OpenContractExporter.cs** - Export documents to OpenContracts format for interoperability:
- `Export(WmlDocument)` / `Export(WordprocessingDocument)` - Export to `OpenContractDocExport`
- Complete text extraction (paragraphs, tables, headers, footers, footnotes, endnotes)
- PAWLS-format page layout with token positions
- Structural annotations (sections, paragraphs, tables) with relationships
- See `docs/architecture/opencontracts_export.md` for detailed documentation

### Target Frameworks

Library targets: `net8.0`
Tests target: `net8.0`

### Dependencies

- **DocumentFormat.OpenXml**: 3.2.0 (Open XML SDK)
- **SkiaSharp**: 2.88.9 (cross-platform graphics, replaces System.Drawing)

### Test Data

Test files are in `TestFiles/` directory with prefixes indicating their purpose:
- `DB*` - DocumentBuilder tests
- `DA*` - DocumentAssembler tests
- `HC*` - HTML Converter tests
- `WC/` - WmlComparer tests
- `SH*` - Spreadsheet tests
- `CU*` - Chart Updater tests

## Migration Status (November 2025)

### Completed

1. **Framework Migration**: Upgraded from net45/net46/netstandard2.0 to .NET 8.0
2. **Open XML SDK 3.x**: Upgraded from 2.8.1 to 3.2.0
   - Replaced `.Close()` with `Dispose()` pattern
   - Added `GetPackage()` extension in `PtOpenXmlUtil.cs` for internal Package access (via reflection)
   - Changed `FontPartType`/`ImagePartType` to `PartTypeInfo` pattern
3. **SkiaSharp Migration**: Replaced System.Drawing with SkiaSharp 2.88.9
   - `SKColor` replaces `Color`
   - `SKBitmap` replaces `Bitmap`
   - `SKFontManager`/`SKTypeface` replaces `FontFamily`/`FontStyle`
   - `SKEncodedImageFormat` replaces `ImageFormat`
   - Created `SkiaSharpHelpers.cs` with `ColorHelper` class for color name mapping
   - Added `SkiaSharp.NativeAssets.Linux.NoDependencies` for Linux runtime support
4. **Test Project**: Updated to .NET 8.0, fixed SkiaSharp usage
5. **WmlComparer Fixes**: Fixed null Unid attribute handling that caused "Internal error" exceptions
6. **Rebranding**: Renamed library from OpenXmlPowerTools to Docxodus
   - Renamed all namespaces to `Docxodus`
   - Renamed `OpenXmlPowerToolsDocument` to `DocxodusDocument`
   - Renamed `OpenXmlPowerToolsException` to `DocxodusException`
   - Archived example projects to `archived-examples/`

### Current Test Status

- **995 passed**, 0 failed, 1 skipped out of 996 tests (~99.9% pass rate)

### Fixed Test Failures (18 tests fixed)

1. **DocumentBuilder relationship tests** (10 tests) - Fixed bug where relationship IDs from source documents could incorrectly match existing IDs in target parts, causing "relationship not found" validation errors
2. **SpreadsheetWriter date handling** (1 test) - Fixed date values being written as ISO 8601 strings instead of Excel serial date numbers
3. **WmlComparer footnote/endnote tests** (6 tests: WC-1660, WC-1670, WC-1710, WC-1720, WC-1750, WC-1760) - Fixed `AssignUnidToAllElements` to assign Unid to footnote/endnote elements themselves, enabling proper reconstruction of multi-paragraph footnotes/endnotes
4. **WmlComparer table row comparison** (1 test: WC-1500) - Added LCS-based row matching for large tables (7+ rows) when content differs significantly, preventing cascading false differences from insertions/deletions in the middle of tables

### Remaining Work

1. **Phase 4**: Remove preprocessor directives (`NET35`, `ELIDE_XUNIT_TESTS`) from source and test files
2. **Phase 6**: Final cleanup and documentation

### Key Files Changed

- `Docxodus.csproj` - Framework and dependency updates
- `Docxodus.Tests.csproj` - Test framework updates
- `PtOpenXmlUtil.cs` - Added `GetPackage()` extension method with SDK 3.x reflection workaround
- `SkiaSharpHelpers.cs` - New file with color utilities
- `ColorParser.cs`, `HtmlToWmlCssParser.cs` - SKColor migration
- `MetricsGetter.cs`, `WmlToHtmlConverter.cs`, `HtmlToWmlConverterCore.cs` - Font/image handling
- `WmlComparer.cs` - Fixed null Unid handling, Package access fixes, footnote/endnote Unid assignment, LCS-based table row matching
- `PresentationBuilder.cs` - Package access fixes
- `DocumentBuilder.cs` - Fixed relationship copying bugs in `CopyRelatedImage`, `CopyRelatedPartsForContentParts`, and related functions
- `SpreadsheetWriter.cs` - Fixed date cell handling to use Excel serial date format

## OOXML Corner Cases

When investigating bugs where our output differs from Word/LibreOffice rendering, **always document findings** in `docs/ooxml_corner_cases.md`. This is critical because:

1. **Word doesn't always follow the spec** - Microsoft Word sometimes implements undocumented behavior or interprets ambiguous spec sections differently than expected
2. **Future reference** - These edge cases are hard to rediscover; documenting them saves hours of debugging later
3. **Test coverage** - Each documented case should eventually have a corresponding test

### What to Document

- Any case where Word renders differently than a literal reading of the OOXML spec would suggest
- Behaviors that differ between Word, LibreOffice, and our implementation
- Numbering/list formatting edge cases (especially legal numbering, multi-level formats)
- Style inheritance quirks
- Table layout anomalies
- Character/paragraph property interactions

### Documentation Format

For each corner case, include:
1. **Minimal XML reproducer** - The smallest XML snippet that demonstrates the issue
2. **Renderer comparison table** - What Word, LibreOffice, and Docxodus each produce
3. **Analysis** - Your hypothesis about why the difference exists
4. **Relevant code** - Which Docxodus files/functions are involved
5. **Proposed fix** - If known, how to align with Word's behavior
