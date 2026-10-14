# docx2html

Convert Word documents (DOCX) to HTML.

## Installation

```bash
dotnet tool install -g Docx2Html
```

## Usage

```bash
# Basic usage - outputs to document.html
docx2html document.docx

# Specify output file
docx2html document.docx output.html

# With custom page title
docx2html document.docx --title="My Document"

# Extract images to separate files instead of embedding as base64
docx2html document.docx --extract-images

# Use inline styles instead of CSS classes
docx2html document.docx --inline-styles
```

## Options

| Option | Description |
|--------|-------------|
| `--title=<text>` | Page title (default: document title or filename) |
| `--css-prefix=<text>` | CSS class prefix (default: pt-) |
| `--inline-styles` | Use inline styles instead of CSS classes |
| `--extract-images` | Save images to separate files instead of embedding |
| `-h, --help` | Show help message |
| `-v, --version` | Show version information |

## Features

- Converts paragraphs, headings, lists, tables, and formatting
- Handles images (embedded as base64 or extracted to files)
- Preserves hyperlinks and bookmarks
- Supports bidirectional (RTL) text
- Generates clean, semantic HTML5

## Limitations

- Math equations (OMML) are not converted
- Charts and diagrams appear as images (if embedded) or are omitted
- Complex text boxes may not render perfectly

## Environment Variables

- `DOCX2HTML_DEBUG=1`: Show detailed error information including stack traces

## License

MIT License
