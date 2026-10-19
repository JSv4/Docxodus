# Bookmark name conflicts

Original literal text and OOXML authored for #1212; MIT, John Scrudato IV. Generated from
`bookmark_conflict_cases()` in JSv4/docx-fidelity-bench, with no imported document content.

Each case has base (`-0`), two reviewer (`-1`, `-2`) inputs and the combined `-final` expectation.
The base is also the rejected expectation. Three conflicts assign one new name to different
block positions, mixed block/inline positions or different inline extents. The control agrees
on the range. Default BaseWins removes conflicting new ranges; an independent text insertion
survives in all four cases.
