# docx2oc

A command-line tool to export Word documents (.docx) to the OpenContracts format.

## Installation

### As a .NET Tool (recommended)

```bash
dotnet tool install --global Docx2OC
```

### Build from Source

```bash
cd tools/docx2oc
dotnet build
```

## Usage

```bash
# Export with default output filename (input.oc)
docx2oc contract.docx

# Export with custom output filename
docx2oc contract.docx export.json
```

## Output Format

The output is a JSON file in the OpenContracts format, containing:

- **title**: Document title from core properties
- **content**: Complete extracted text (body, headers, footers, footnotes, endnotes)
- **description**: Document description/subject if available
- **pageCount**: Estimated page count
- **pawlsFileContent**: PAWLS-format page layout with token positions
- **labelledText**: Structural annotations (SECTION, PARAGRAPH, TABLE)
- **relationships**: Hierarchical relationships between annotations

## Example Output

```json
{
  "title": "Sample Contract",
  "content": "This is the document content...",
  "pageCount": 5,
  "pawlsFileContent": [
    {
      "page": { "width": 612, "height": 792, "index": 0 },
      "tokens": [
        { "x": 72, "y": 72, "width": 30, "height": 12, "text": "This" }
      ]
    }
  ],
  "labelledText": [
    {
      "id": "section-0",
      "annotationLabel": "SECTION",
      "structural": true
    }
  ]
}
```

## OpenContracts Compatibility

This tool produces output compatible with the [OpenContracts](https://github.com/JSv4/OpenContracts) document analysis platform. The format includes:

- Complete text extraction from all document parts
- PAWLS-compatible token positions for NLP/ML pipelines
- Structural annotations for document understanding

## Environment Variables

- `DOCX2OC_DEBUG=1`: Show detailed error information including stack traces
