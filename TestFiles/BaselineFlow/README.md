# Native Word baseline pitch

The eight long paragraphs in `baseline-rounding-font-size-holdouts.docx` cover Times New Roman
and Arial at 10, 13 and 16 points with 1.15 auto spacing, plus 12-point Times New Roman controls
at single and 1.5 spacing. Each paragraph sets both its paragraph-mark and text-run font and size;
the default paragraph mark cannot supply a larger hidden line box. The input has zero Open XML
SDK validation findings.

`word-observations.json` records the exact source and native PDF hashes, native page membership,
line text and baselines from an unedited Word for the web PDF export on 2026-10-25. Word's
reported build was unavailable. The original DOCX, returned Word DOCX and raw native PDF are
preserved in the issue's evidence archive.

The browser regression checks relative baseline pitch within each uninterrupted page segment.
Its open-font scaffolds use Liberation Serif and Liberation Sans, whose relevant vertical metrics
match the captured fonts; public-export diagnostics use the exact declared installed fonts.
The assertions retain the captured PDF placement variation and reject cumulative loss from
rounding before applying the spacing multiple. They do not establish arbitrary raster parity or
close the agreement's remaining justification issue.
