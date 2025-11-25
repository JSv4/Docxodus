# Redline

A command-line tool for comparing Word documents and generating redline diffs with tracked changes.

## Installation

### As a .NET Global Tool

```bash
dotnet tool install -g Redline
```

### From GitHub Packages

```bash
dotnet nuget add source https://nuget.pkg.github.com/jmansdorff/index.json --name github
dotnet tool install -g Redline --source github
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

## Powered By

This tool is built on [OpenXmlPowerTools](https://github.com/jmansdorff/Open-Xml-PowerTools), specifically the `WmlComparer` module for document comparison.

## License

MIT License - see [LICENSE](https://github.com/jmansdorff/Open-Xml-PowerTools/blob/main/LICENSE) for details.
