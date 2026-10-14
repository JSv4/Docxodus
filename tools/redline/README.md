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
