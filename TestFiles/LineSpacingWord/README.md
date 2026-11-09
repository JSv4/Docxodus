# Original Word line-spacing references

These two small, synthetic DOCX fixtures were supplied by the issue author in the public
#1565/#1566 discussions. `word-observations.json` records the comment URLs, exact input hashes,
unedited Word Online PDF hashes, visible lines, font metrics and baselines. The PDFs were
exported on 2026-11-09 UTC using File > Export > Download as PDF. The first download was performed
by the user; the second through the signed-in OneDrive editor. No document text was edited.

The 16pt-on 10pt fixture puts its first baseline 85.78003pt below the page top, whose margin is 72pt.
Its natural descent is 443/2048 of 10pt. This supports a baseline at line height minus descent
(85.83691pt), rejecting centered leading (83.58643pt). The mixed 16/9pt1.15 fixture supports
automatic surplus below the baseline: Word's first transition is 14.75pt versus 14.62305pt from
that model. These are distinct policies. Trailing 12pt PDF spaces are absent from the authored
DOCX text and must not be introduced as new document runs.

Word embeds Times New Roman 7.06. Local comparison used installed Times New Roman 2.82; their
hhea metrics and the 43 characters used by the mixed fixture have matching advances. Portable
regressions use Liberation Serif, which shares the relevant vertical metrics. No standalone
Microsoft font file is included. The fixtures are original synthetic prose supplied for this
repository's regression work; the PDFs retain Word's embedded subsets.

To refresh, upload each unchanged DOCX to Word, open it from OneDrive, choose the PDF export,
then confirm Download in the additional dialog. Preserve the returned PDF bytes and extract
text origins in points from the page top. Update hashes and observations together. This evidence
does not establish blank-break, marker-transition or mixed-minimum parity.
