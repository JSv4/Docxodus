# Move range membership fixtures

Six original literal-OOXML documents, authored by John Scrudato IV under MIT, pin issue #1340.
They contain no external document content or embedded fonts. Input bytes are immutable.

`word-endpoints.json` records the input hashes and paragraph lists observed after native Word for
the web accept-all/reject-all on 2026-10-24. Each endpoint has native DOCX, document XML and PDF
hashes. Word did not expose an exact build. Native downloads were preserved unchanged, including
unrelated settings-order validator findings; only the source inputs are used as schema-clean tests.
The evidence ZIP hash identifies the complete capture with UI proof and rendered previews.

Controls are `unsplit-control` and `plain-outside-control`; the other four inputs exercise live text
inside a range or multiple disjoint destination ranges with the same move name. The tests compare
all body paragraphs, require no surviving move markers and validate every source/result against
Open XML SDK 3.5.1, Microsoft365. Unit tests separately cover the mirrored source side, moved
paragraph marks, overlapping ranges, legacy Word-authored moves and non-run metadata preservation.
