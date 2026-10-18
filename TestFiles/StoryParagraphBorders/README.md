These four original fixtures were authored for issue #1203 by
[`bench/fixtures.py`](https://github.com/JSv4/docx-fidelity-bench/blob/fixtures/1203-story-paragraph-borders/bench/fixtures.py)
in JSv4/docx-fidelity-bench (MIT, John Scrudato IV). No source document, template,
external text, image or embedded font was used. The files are byte-identical to
that benchmark's `paragraphs-story-borders-*` PDF-track inputs.

The native and WASM/export tests share these bytes. Each document includes body,
header, footer and note cards: two adjacent paragraphs with equal borders followed
by one plain paragraph. Each unbordered control differs only in paragraph borders.
The note kind is either footnote or endnote. The requested font is Liberation Serif.
