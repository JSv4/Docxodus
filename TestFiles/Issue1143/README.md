Comparison pair for issue #1143 (a one-word edit inside an unchanged block content control).

Both files derive from `TestFiles/DB006-Source1.docx`, a Word-authored article, by one XML edit
to `word/document.xml`; every other part is byte-identical to the source.

- `clause-control-original.docx`: the body-level run from the "Block-Level Content Containers"
  heading through the paragraph beginning "As I mentioned" (a heading, a prose paragraph, the
  seven-row table, and a prose paragraph that carries a `w:proofErr` pair and a
  `w:lastRenderedPageBreak`) is wrapped in one rich-text block control
  (`w:alias` "Block-level containers", `w:tag` `clause-containers`, `w:id` 1476395012).
- `clause-control-revised.docx`: the same, with "There are **only** three block-level content
  container elements" changed to "There are **just** three" inside the control.

`BlockWrapperNestedDiffTests` compares the pair through the `DocxCompare` front door and checks
that the control is kept once, that the redline marks exactly that word, that accept and reject
reproduce the revised and original documents block for block, and that the redline adds no
Office 2019 validation error.
