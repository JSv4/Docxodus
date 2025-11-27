# Redline

**Compare Word documents and generate redlines with tracked changes.**

[![CI](https://github.com/JSv4/DocxRedlines/actions/workflows/ci.yml/badge.svg)](https://github.com/JSv4/DocxRedlines/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

Redline is a .NET tool for comparing two Word documents (DOCX) and producing a third document with tracked changes showing insertions, deletions, and modifications.

## Quick Start

### Install the CLI Tool

```bash
# Add GitHub Packages source (one-time setup)
dotnet nuget add source https://nuget.pkg.github.com/JSv4/index.json \
  --name github \
  --username YOUR_GITHUB_USERNAME \
  --password YOUR_GITHUB_PAT

# Install globally
dotnet tool install -g Redline --source github
```

### Usage

```bash
redline original.docx modified.docx output.docx
```

With a custom author tag for tracked changes:

```bash
redline original.docx modified.docx output.docx --author="Legal Review"
```

### Options

| Option | Description |
|--------|-------------|
| `--author=<name>` | Author name for tracked changes (default: "Redline") |
| `-h, --help` | Show help message |
| `-v, --version` | Show version information |

## DOCX to HTML Conversion

### Install the docx2html CLI Tool

```bash
# Install globally (after adding GitHub Packages source)
dotnet tool install -g Docx2Html --source github
```

### Usage

```bash
# Basic conversion
docx2html document.docx

# Specify output file
docx2html document.docx output.html

# Extract images to files instead of embedding as base64
docx2html document.docx --extract-images

# Use inline styles instead of CSS classes
docx2html document.docx --inline-styles
```

### Options

| Option | Description |
|--------|-------------|
| `--title=<text>` | Page title (default: document title or filename) |
| `--css-prefix=<text>` | CSS class prefix (default: "pt-") |
| `--inline-styles` | Use inline styles instead of CSS classes |
| `--extract-images` | Save images to separate files instead of embedding |
| `-h, --help` | Show help message |
| `-v, --version` | Show version information |

## Using as a Library

Reference the OpenXmlPowerTools project directly in your solution:

```csharp
using OpenXmlPowerTools;

// Load documents
var original = new WmlDocument("original.docx");
var modified = new WmlDocument("modified.docx");

// Configure comparison
var settings = new WmlComparerSettings
{
    AuthorForRevisions = "Redline",
    DetailThreshold = 0
};

// Compare and get result
var result = WmlComparer.Compare(original, modified, settings);

// Get list of revisions
var revisions = WmlComparer.GetRevisions(result, settings);
Console.WriteLine($"Found {revisions.Count} revisions");

// Save the redlined document
result.SaveAs("redline.docx");
```

## Download Standalone Binaries

Pre-built binaries are available on the [Releases](https://github.com/JSv4/DocxRedlines/releases) page:

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
git clone https://github.com/JSv4/DocxRedlines.git
cd DocxRedlines

# Build
dotnet build

# Run tests
dotnet test

# Run the CLI
dotnet run --project tools/redline/redline.csproj -- --help
```

## Project Status

This project is a focused fork of [Open-Xml-PowerTools](https://github.com/OfficeDev/Open-Xml-PowerTools), originally developed by Eric White at Microsoft. The original project provided a broad set of utilities for working with Office Open XML documents but is no longer actively maintained.

**Redline** narrows the focus to document comparison—the most commonly needed feature for legal, editorial, and business workflows.

### What's Included

- **WmlComparer** - Compare two DOCX files and generate redlines with tracked changes
- **WmlToHtmlConverter** - Convert DOCX files to HTML with CSS styling
- **redline** CLI tool - Command-line interface for document comparison
- **docx2html** CLI tool - Command-line interface for HTML conversion
- Supporting utilities for document manipulation

## Requirements

- .NET 8.0 or later

## License

MIT License - see [LICENSE](LICENSE) for details.

---

*Built on the shoulders of [Open-Xml-PowerTools](https://github.com/OfficeDev/Open-Xml-PowerTools). Thanks to Eric White, Thomas Barnekow, and all original contributors.*
