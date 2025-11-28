# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Important Instructions

- **Never credit yourself in commits.** Do not add "Generated with Claude Code" or "Co-Authored-By: Claude" to commit messages.

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

**WmlComparer.cs** - Compare two DOCX files, producing a document with tracked revisions. Supports nested tables and text boxes. Key settings in `WmlComparerSettings`.

**WmlToHtmlConverter.cs / HtmlToWmlConverter.cs** - Bidirectional DOCX ↔ HTML conversion. Supports tracked changes rendering via `WmlToHtmlConverterSettings`:
- `RenderTrackedChanges` - Render insertions/deletions as `<ins>`/`<del>` instead of accepting them
- `RenderMoveOperations` - Distinguish move operations from regular insert/delete
- `RenderFootnotesAndEndnotes` - Include footnotes/endnotes sections in HTML output
- `RenderHeadersAndFooters` - Include document headers/footers in HTML output
- `AuthorColors` - Dictionary mapping author names to CSS colors for styling

**DocumentAssembler.cs** - Template population from XML data using content controls.

**PresentationBuilder.cs** - Merge/split PPTX files.

**SpreadsheetWriter.cs** - Simplified XLSX creation API with streaming support for large files.

**OpenXmlRegex.cs** - Search/replace in DOCX/PPTX using regular expressions.

**RevisionAccepter.cs / RevisionProcessor.cs** - Handle tracked revisions.

**FormattingAssembler.cs** - Resolve and flatten document formatting.

**MetricsGetter.cs** - Extract document metrics (styles, fonts, languages).

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

- **989 passed**, 0 failed, 1 skipped out of 990 tests (~99.9% pass rate)

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
