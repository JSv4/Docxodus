# Redline

A command-line tool for comparing Word documents and generating redlines with tracked changes.

## Installation

### As a .NET Global Tool

```bash
dotnet tool install --global Redline
```

## Usage

```bash
redline <original.docx> <modified.docx> <output.docx> [--author=<name>]
```

### Arguments

| Argument | Description |
|----------|-------------|
| `original.docx` | Path to the original document |
| `modified.docx` | Path to the modified document |
| `output.docx` | Path for the output redline document |

### Options

| Option | Description |
|--------|-------------|
| `--author=<name>` | Author name for tracked changes (default: "Redline") |
| `--minimum-anchor-word-count=<n>` | How many words a paragraph needs before the comparison uses it to line the two documents up (default: 0). Use 2 for documents whose entries each start with a one-word number label on its own line; see the example below. |
| `-h, --help` | Show help message |
| `-v, --version` | Show version information |

## Examples

Basic comparison:
```bash
redline contract-v1.docx contract-v2.docx redline.docx
```

With custom author tag:
```bash
redline draft.docx final.docx changes.docx --author="Legal Review"
```

Numbered entries. Some documents put a short label on its own line before each entry: "P1", then the
entry's text, then "P2", and so on. Adding or removing an entry changes every later label, and the
comparison can then report the entries' unchanged text as moved. Requiring two words before a paragraph
is used to line the documents up makes it match the entries by their text instead, so each label just
shows its number changing:
```bash
redline schedule-v1.docx schedule-v2.docx changes.docx --minimum-anchor-word-count=2
```

## Output

The tool generates a Word document with tracked changes (revisions) showing:
- **Insertions**: Text added in the modified document
- **Deletions**: Text removed from the original document
- **Formatting changes**: Style and formatting differences

Open the output document in Microsoft Word or another compatible word processor to review and accept/reject changes.

## Environment Variables

- `REDLINE_DEBUG=1`: Show detailed error information including stack traces

## License

MIT License - see [LICENSE](https://github.com/JSv4/Docxodus/blob/main/LICENSE) for details.
