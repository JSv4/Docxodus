# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Important Instructions

- **Never credit yourself in commits.** Do not add "Generated with Claude Code" or "Co-Authored-By: Claude" to commit messages.

## Build Commands

```bash
# Build the entire solution
dotnet build OpenXmlPowerTools.sln

# Build specific project
dotnet build OpenXmlPowerTools/OpenXmlPowerTools.csproj
```

## Test Commands

```bash
# Run all tests
dotnet test OpenXmlPowerTools.Tests/OpenXmlPowerTools.Tests.csproj

# Run a specific test by name
dotnet test --filter "FullyQualifiedName~DB001_DocumentBuilderKeepSections"

# Run tests for a specific test class
dotnet test --filter "FullyQualifiedName~DbTests"
```

## Architecture Overview

Open-Xml-PowerTools is a library for manipulating Open XML documents (DOCX, XLSX, PPTX) built on top of the Open XML SDK. All code is in the `OpenXmlPowerTools` namespace.

### Document Wrapper Classes

The library uses in-memory byte array wrappers for documents:
- `OpenXmlPowerToolsDocument` - Base class holding `DocumentByteArray` and `FileName`
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

**WmlToHtmlConverter.cs / HtmlToWmlConverter.cs** - Bidirectional DOCX ↔ HTML conversion.

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

### Current Test Status

- **951 passed**, 27 failed out of 979 tests (~97% pass rate)

### Remaining Test Failures (27)

1. **WmlComparer footnote/endnote tests** (7 failures) - Some comparison tests return 0 revisions for documents with footnotes/endnotes/tables. May require deeper investigation into how Unid attributes are being generated/compared.

2. **Cell format currency tests** (7 failures) - Currency formatting produces slightly different output (e.g., "-₩ -" instead of "₩ -"). Likely pre-existing issue.

3. **DocumentBuilder relationship tests** (9 failures) - Errors like "The relationship 'rId11' referenced by attribute 'r:embed' does not exist." Likely related to SDK 3.x changes in relationship handling.

4. **Other tests** (4 failures) - PowerToolsBlockExtensionsTests, PB006_VideoFormats, SW002_AllDataTypes

### Remaining Work

1. **Phase 4**: Remove preprocessor directives (`NET35`, `ELIDE_XUNIT_TESTS`) from source and test files
2. **Phase 5**: Update example projects (6 example projects still target old frameworks)
3. **Phase 6**: Investigate and fix remaining 27 test failures
4. **Phase 7**: Final cleanup and documentation

### Key Files Changed

- `OpenXmlPowerTools.csproj` - Framework and dependency updates
- `OpenXmlPowerTools.Tests.csproj` - Test framework updates
- `PtOpenXmlUtil.cs` - Added `GetPackage()` extension method with SDK 3.x reflection workaround
- `SkiaSharpHelpers.cs` - New file with color utilities
- `ColorParser.cs`, `HtmlToWmlCssParser.cs` - SKColor migration
- `MetricsGetter.cs`, `WmlToHtmlConverter.cs`, `HtmlToWmlConverterCore.cs` - Font/image handling
- `WmlComparer.cs` - Fixed null Unid handling, Package access fixes
- `PresentationBuilder.cs` - Package access fixes
