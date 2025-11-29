<p align="center">
  <img src="docxodus-mono-final.svg" alt="Docxodus" width="400">
</p>

<p align="center">
  <strong>A powerful .NET library for manipulating Open XML documents (DOCX, XLSX, PPTX).</strong>
</p>

<p align="center">
  <a href="https://github.com/JSv4/Redlines/actions/workflows/ci.yml"><img src="https://github.com/JSv4/Redlines/actions/workflows/ci.yml/badge.svg" alt="CI"></a>
  <a href="https://opensource.org/licenses/MIT"><img src="https://img.shields.io/badge/License-MIT-yellow.svg" alt="License: MIT"></a>
</p>

---

Docxodus is a fork of [Open-Xml-PowerTools](https://github.com/OfficeDev/Open-Xml-PowerTools) upgraded to .NET 8.0. It provides tools for comparing Word documents, converting between DOCX and HTML, merging documents, and more.

## Quick Start

### Install the Library

```bash
# Add GitHub Packages source (one-time setup)
dotnet nuget add source https://nuget.pkg.github.com/JSv4/index.json \
  --name github \
  --username YOUR_GITHUB_USERNAME \
  --password YOUR_GITHUB_PAT

# Add to your project
dotnet add package Docxodus --source github
```

### Using as a Library

```csharp
using Docxodus;

// Compare documents
var original = new WmlDocument("original.docx");
var modified = new WmlDocument("modified.docx");

var settings = new WmlComparerSettings
{
    AuthorForRevisions = "Redline",
    DetailThreshold = 0
};

var result = WmlComparer.Compare(original, modified, settings);

// Get list of revisions (with move detection)
var revisions = WmlComparer.GetRevisions(result, settings);
foreach (var rev in revisions)
{
    if (rev.RevisionType == WmlComparer.WmlComparerRevisionType.Moved)
        Console.WriteLine($"Moved (group {rev.MoveGroupId}): {rev.Text}");
    else
        Console.WriteLine($"{rev.RevisionType}: {rev.Text}");
}

// Save the redlined document
result.SaveAs("redline.docx");
```

## CLI Tools

Docxodus includes two command-line tools:

### Redline (Document Comparison)

```bash
# Install globally
dotnet tool install -g Redline --source github

# Usage
redline original.docx modified.docx output.docx

# With custom author tag
redline original.docx modified.docx output.docx --author="Legal Review"
```

| Option | Description |
|--------|-------------|
| `--author=<name>` | Author name for tracked changes (default: "Redline") |
| `-h, --help` | Show help message |
| `-v, --version` | Show version information |

### docx2html (HTML Conversion)

```bash
# Install globally
dotnet tool install -g Docx2Html --source github

# Basic conversion
docx2html document.docx

# Specify output file
docx2html document.docx output.html

# Extract images to files instead of embedding as base64
docx2html document.docx --extract-images

# Use inline styles instead of CSS classes
docx2html document.docx --inline-styles
```

| Option | Description |
|--------|-------------|
| `--title=<text>` | Page title (default: document title or filename) |
| `--css-prefix=<text>` | CSS class prefix (default: "pt-") |
| `--inline-styles` | Use inline styles instead of CSS classes |
| `--extract-images` | Save images to separate files instead of embedding |
| `-h, --help` | Show help message |
| `-v, --version` | Show version information |

## Download Standalone Binaries

Pre-built binaries are available on the [Releases](https://github.com/JSv4/Redlines/releases) page:

**redline** (Document Comparison):

| Platform | Download |
|----------|----------|
| Windows (x64) | `redline-win-x64.exe` |
| Linux (x64) | `redline-linux-x64` |
| macOS (x64) | `redline-osx-x64` |
| macOS (ARM) | `redline-osx-arm64` |

**docx2html** (HTML Conversion):

| Platform | Download |
|----------|----------|
| Windows (x64) | `docx2html-win-x64.exe` |
| Linux (x64) | `docx2html-linux-x64` |
| macOS (x64) | `docx2html-osx-x64` |
| macOS (ARM) | `docx2html-osx-arm64` |

## Build from Source

```bash
# Clone the repository
git clone https://github.com/JSv4/Redlines.git
cd Redlines

# Build
dotnet build Docxodus.sln

# Run tests
dotnet test Docxodus.Tests/Docxodus.Tests.csproj

# Run the CLI
dotnet run --project tools/redline/redline.csproj -- --help
```

## Features

- **WmlComparer** - Compare two DOCX files and generate redlines with tracked changes
  - **Move Detection** - Automatically detects when content is relocated (not just deleted and re-inserted)
  - Configurable similarity threshold and minimum word count
  - Links move pairs via `MoveGroupId` for easy tracking
- **WmlToHtmlConverter** / **HtmlToWmlConverter** - Bidirectional DOCX ↔ HTML conversion
- **DocumentBuilder** - Merge and split DOCX files
- **DocumentAssembler** - Template population from XML data
- **PresentationBuilder** - Merge and split PPTX files
- **SpreadsheetWriter** - Simplified XLSX creation API
- **OpenXmlRegex** - Search/replace in DOCX/PPTX using regular expressions
- Supporting utilities for document manipulation

## Browser/JavaScript Usage (npm)

Docxodus is also available as an npm package for client-side usage via WebAssembly:

```bash
npm install docxodus
```

```javascript
import {
  initialize,
  convertDocxToHtml,
  compareDocuments,
  getRevisions,
  isMove,
  isMoveSource,
  findMovePair,
  CommentRenderMode
} from 'docxodus';

await initialize();

// Convert DOCX to HTML with comments
const html = await convertDocxToHtml(docxFile, {
  commentRenderMode: CommentRenderMode.EndnoteStyle
});

// Compare two documents
const redlinedDocx = await compareDocuments(originalFile, modifiedFile);

// Get revisions with move detection
const revisions = await getRevisions(redlinedDocx);
for (const rev of revisions.filter(isMove)) {
  const pair = findMovePair(rev, revisions);
  if (isMoveSource(rev)) {
    console.log(`Content moved from: "${rev.text}" to: "${pair?.text}"`);
  }
}
```

See the [npm package documentation](docs/npm-package.md) for full API reference, React hooks, and usage examples.

## Requirements

- .NET 8.0 or later

## License

MIT License - see [LICENSE](LICENSE) for details.

---

*Built on the shoulders of [Open-Xml-PowerTools](https://github.com/OfficeDev/Open-Xml-PowerTools). Thanks to Eric White, Thomas Barnekow, and all original contributors.*
