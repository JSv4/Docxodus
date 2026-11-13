# Metafile fixtures

Fixtures for `Docxodus.Tests/MetafileRendererTests.cs` (issue #1706).

| File | What it is |
| --- | --- |
| `libreoffice-chart.source.svg` | The source drawing, written for this repository by the contributor of PR #1713. It holds no third-party content. |
| `libreoffice-chart.emf` | LibreOffice 26.8 export of the source drawing (`soffice --headless --convert-to emf`). A dual EMF+ file: its GDI records carry paths, Béziers, text ("Quarterly 42", Liberation Sans), a clip and one 2×2 STRETCHDIBITS bitmap. |
| `libreoffice-chart.wmf` | The same drawing exported with `--convert-to wmf`. LibreOffice embeds a copy of the EMF in two `META_ESCAPE` records; that copy is byte-identical to `libreoffice-chart.emf`. |

The files are covered by the repository's licence like the rest of `TestFiles/`. Both exports end at their end-of-file
record with no trailing bytes.

Tests for malformed or hostile records build their bytes in the test itself; do not add crafted binary files here.
