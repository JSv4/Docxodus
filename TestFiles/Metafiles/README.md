# Metafile fixtures

Fixtures for `Docxodus.Tests/MetafileRendererTests.cs` (issue #1706).

| File | What it is |
| --- | --- |
| `libreoffice-chart.source.svg` | The source drawing, written for this repository by the contributor of PR #1713. It holds no third-party content. |
| `libreoffice-chart.emf` | LibreOffice 26.8 export of the source drawing (`soffice --headless --convert-to emf`). A dual EMF+ file: its GDI records carry paths, Béziers, text ("Quarterly 42", Liberation Sans), a clip and one 2×2 STRETCHDIBITS bitmap. |
| `libreoffice-chart.wmf` | The same drawing exported with `--convert-to wmf`. LibreOffice embeds a copy of the EMF in two `META_ESCAPE` records; that copy is byte-identical to `libreoffice-chart.emf`. |

The `office-*` files were written on Windows 11 by Office 2024 (16.0) through its COM automation, from content made
for this repository: the charts use the sample data Office inserts into a new chart, and every shape and string was
typed for the purpose. Each was read string by string before it was added: the only text in them is font names and
the text listed below, with no user, machine or path names. They are the bytes Office wrote, unchanged.

| File | What it is |
| --- | --- |
| `office-scatter-markers.emf` | Word's `Range.EnhMetaFileBits` of a scatter chart with markers titled "Sales scatter-markers". A dual EMF+ file whose GDI records draw each marker as a pair of STRETCHDIBITS between two `TNPP` comments: a one-bit mask with SRCPAINT, then the image with SRCAND. |
| `office-area-gradient.emf` | PowerPoint's `Shape.Export` (EMF) of an area chart titled "Sales area-gradient" whose chart area and first series have two-colour gradient fills. Dual EMF+; the gradients are strips of solid polygons, and its clip paths are drawn inside BEGINPATH under a 1/8 world scale that is undone before ENDPATH. |
| `office-symbols.emf` | Word's `Range.EnhMetaFileBits` of two paragraphs: Symbol and Wingdings characters inserted as symbols, then Unicode symbols in Calibri. The symbol fonts' characters are stored as their codes moved to U+F020–U+F0FF. |
| `office-symbols-converted.wmf` | `office-symbols.emf` converted by GDI's `GetWinMetaFileBits` (MM_ANISOTROPIC, the conversion Windows makes for a "Picture (Windows Metafile)" clipboard format), with a placeable header added from its window extent. Its ANSI text records lose the Greek and arrow characters; it also carries `office-symbols.emf`, byte for byte, in META_ESCAPE_ENHANCED_METAFILE comments. |
| `office-pattern-star.emf`, `office-pattern-star.wmf` | PowerPoint's `Shape.Export` (EMF and WMF) of a five-point star filled with a red-on-white diagonal pattern: an 8×8 one-bit DIB pattern brush in RGB colours. |
| `office-pie-chart.wmf` | PowerPoint's `Shape.Export` (WMF) of a pie chart with data labels titled "Sales pie-labels". It resets the clip with SELECTCLIPREGION 0 before each INTERSECTCLIPRECT. |
| `office-dashed-arrow.wmf` | PowerPoint's `Shape.Export` (WMF) of a green dashed line ending in an arrowhead, whose outline is drawn with a zero-width pen. |

The files are covered by the repository's licence like the rest of `TestFiles/`. Both LibreOffice exports end at their
end-of-file record with no trailing bytes.

Tests for malformed or hostile records build their bytes in the test itself; do not add crafted binary files here.
