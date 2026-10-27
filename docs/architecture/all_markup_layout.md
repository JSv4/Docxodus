# All Markup print layout (issue #1300)

**Status: implemented.** The values below were measured from Word-printed All Markup PDFs, except
where marked *pending*.

## What it does

When Word prints or saves a document as PDF with All Markup and "Show only comments and formatting in
balloons", it keeps each page's paper size and shrinks the page so that it fits beside a grey markup
area on its right. Comments, their replies and formatting changes become balloons in that area, level
with their lines and joined to them by leader lines. Insertions and deletions stay inline.

Docxodus produces the same layout when a paginated render uses `commentColumnPlacement: "balloons"`.
It is available in the paginated view (`paginateHtml`, `PaginationEngine`, the React viewer) and in the
`@docxodus/export` HTML and PDF (`commentColumnPlacement`, or `--comment-column-placement balloons` in
the CLI). With the option off, output does not change.

## Option

`commentColumnPlacement` already decides where a page's comment column goes: `"margin"` inside the
page's right margin, `"outside"` in a panel beside the page box. `"balloons"` is a third answer to the
same question, so the layout extends that option rather than adding a new one. Formatting balloons
share the markup area because Word ties them to the same setting.

Alternatives that were rejected:

- **A new `commentProfile` value.** Comment profiles choose how the converter renders comments, and
  delivery bundles share their strings. Margin mode's converter output already carries the comment
  balloon data.
- **A boolean beside `commentProfile`.** That would mean two options describing one geometry.
- **A new converter setting.** It would ripple through every transport, for data that only the
  paginator reads.

The export accepts `"margin" | "balloons"` (`"outside"` widens the output past the paper), and
validates them as follows:

- `"balloons"` takes any comment profile but `inline`, whose comments sit in the text: `margin`
  comments go in the balloons; `endnotes` comments stay in their list after the text, with their
  reference marks and highlights, and the balloons hold the formatting changes; `hidden` prints
  formatting balloons only.
- With the markup review profile it also needs `markupPresentation: "word"`.

## Responsibilities

### Converter: records, never lays out

Everything below happens only in `RevisionPresentation.Word`. The Docxodus presentation's HTML is
unchanged, and its tooltips (`title`, from `DescribeFormatChange`) keep their wording.

**Formatting wording (`Docxodus/Internal/WordFormatWording.cs`).**

- Before formatting assembly, `Stamp` visits every change in the content parts:
  - `w:rPrChange` on runs and on paragraph marks;
  - `w:pPrChange` on paragraphs.
- It resolves both sides of each change through the assembler's own non-mutating resolvers
  (`FormattingAssembler.ResolveEffectiveRunProperties` / `ResolveEffectiveParagraphProperties`), so
  whatever the styles and document defaults supply equally to both sides cancels out. This is why
  no second style resolver exists.
- It writes its results onto the change element as attributes in the
  `urn:docxodus:rendering:balloon` namespace. Assembly carries change elements through unchanged.
  Runs that repeat a change in the same paragraph context reuse the first result, so long
  documents do not resolve styles once per run.
- At render time the converter publishes every such attribute as `data-` plus its name, on the
  `rev-format-change` span or the `rev-para-format-change` paragraph:
  - `data-format-description`: the wording. An empty value means a bare "Formatted".
  - `data-author-color`: the revision author's resolved colour.
  - `data-format-key` (runs only): an ordinal shared by every change that recorded the same
    properties.
- A list number shows its paragraph mark's formatting and gets no wording of its own.

**Comment balloon data.** Margin notes carry:

- `data-initials` (`w:initials`);
- `data-author-color`: comment authors are numbered on their own, in the order the comments part
  first names them, and `AuthorColors` overrides by name;
- `data-font-family`: the font of the comment's first text run, which its balloon keeps.

**Balloon font.** The pagination staging element carries `data-balloon-font-family` and
`data-balloon-font-size`. They come from the document's Balloon Text paragraph style when it has
one, and otherwise default to Segoe UI 9 pt.

**The Word presentation itself.**

- Its author palette starts with Word's first two author colours, `#0078D4` and `#D13438`. The other
  six (`#5C2E91`, `#038387`, `#CA5010`, `#498205`, `#C239B3`, `#986F0B`) are this implementation's
  choice and were not measured from Word.
- Its change bar sits 30pt left of the text column.

These apply with or without balloons, so they are listed under `### Changed`.

### Paginator: lays out (`npm/src/markup-balloons.ts`)

Before measuring, `paginate()` (and the export, before its own measuring passes) turns on two rules
for comments that are in the balloons:

- comment reference marks take no width, because Word prints none;
- the commented range loses the converter's highlight.

The rules apply only inside elements carrying `data-docx-markup-balloons="comments"`, which
`scopeMarkupBalloons` puts on that pagination's staging and page containers when the converter's
margin-note registry is present. Under the endnotes profile the attribute is empty: the comments keep
their reference marks and highlights, and only the balloon geometry rules apply. A later pagination
with another placement removes the attribute from its containers, so the rules never leak into it.

After the page tree is final, `layoutMarkupBalloons()` runs. It is called at the end of `paginate()`,
or explicitly with `deferMarkupBalloons` after the export's text and border painting. If the
document has no formatting change and no comment that the balloons show (comments listed as endnotes
have anchors, but no balloon), the pages are left as they are. Otherwise:

**Sheet.**

- The page box keeps the paper size and stays the print unit.
- Its content moves into one layer. That layer is the page as far as its text column's right edge
  plus a 266.4pt (3.7in) markup area, shrunk by `paperWidth / (textRight + 266.4pt)` to the paper's
  width.
- The shrunken page stands at the paper's foot, 10.8 paper-points above it.
- The grey area (`#F2F2F2`) starts 9pt right of the text column and spans the page's height.
- The layer is anchored at the page box's foot and reaches above it while unscaled. This is
  because print fragmentation would carry anything laid out past the foot onto a new page.
- `data-docx-markup-scale` on the page box records the factor.

**Balloons.**

- One balloon per comment and per reply, numbered in document order: a root as
  `Commented [AA1]:`, a reply under its root as `Commented [BB2R1]:`, with initials derived from
  the name when `w:initials` is absent.
- One balloon per formatting revision: consecutive changes with the same wording, colour and
  recorded properties share a balloon across list numbers and untracked paragraph marks. A
  tracked paragraph mark between two changes ends the revision. A mark stands at its paragraph's
  end, so the marks between two changes are the earlier change's own paragraph mark and those of
  whole paragraphs between them; the later change's paragraph mark comes after it and does not
  count. A change in a header or footer and one in the body are never one revision.
  Description: `Formatted:` followed by the wording. A paragraph's balloon follows its runs'.
- A comment thread's balloons are anchored where its range ends, on the page and line holding that
  end, since Word's leader starts there. A point comment, an empty-range comment, or a reply with no
  range of its own is anchored at its reference mark's line. A run change is anchored at its first
  line. A paragraph change is anchored at its last line, on the page that holds its mark.
- Changes and comments in a header or footer get balloons on every page that header or footer
  prints on, because it renders there. The margin placement does the same with its notes
  (issue #1310). Non-body comments follow Docxodus's preservation contract; Word ignores these
  anchors, as documented in `comment_rendering.md`, "Comments outside the body".
- Geometry, in page points before shrinking:

  | Property | Value |
  |---|---|
  | Left edge | 31.6pt past the text column |
  | Right inset | 4pt |
  | Border | 0.5pt, in the author's colour |
  | Corner radius | 3pt |
  | Padding | 1.8pt × 4pt |
  | Line height | 1.15 |
  | Vertical gap between balloons | 1pt |
  | Inset from the page's top and bottom edges | 2pt |

- Fills: comment balloons are tinted (`#D5EDFF` for `#0078D4` and `#F8DCDD` for `#D13438`; other
  colours are mixed mostly with white); formatting balloons are white.
- Labels are bold. The label and the wording use the Balloon Text font; a comment's text keeps its
  own font.

**Placement.**

- Balloons are sorted by line, then by document order within a line.
- Each sits level with its anchor's line where room allows, pushed down past the one above, then
  pulled up from the page's foot as far as the room above allows. Balloons never overlap.
- A stack the page's height can hold is never cut, however low its anchors: pulled up, it prints
  whole. If the page cannot hold it even pulled up, Word's rule applies, measured from its Save as
  PDF and Print to PDF of a page with 25 long comments (3 whole, 22 cut, none moved above its line):
  balloons stay whole from the top for as long as the rest, each cut to its first line, still follow
  at or below their own lines and end above the page's foot. The first balloon that cannot stay
  whole, and every one after it, is cut; none is moved above its line to keep another whole. A
  balloon of one line is never cut. Between the two measured cases (a stack that fits only when
  pulled up) the first rule is assumed. A cut balloon is 16.1pt tall at the 9pt balloon font (scaled
  with it) and ends in a box at its lower right corner, 11.6pt tall with 4pt horizontal padding, a
  0.2pt black border and a white fill, holding an ellipsis; the line's text ends at least 1pt before
  the box, and only the first line's own box stays visible. The box is laid out in the balloon's flow
  (the balloon becomes a two-column grid), because positioned absolutely it made Chromium shrink every
  printed page. If even the cut balloons cannot fit,
  they run off the page's foot and are cut there. Either way the balloon is marked
  `data-balloon-truncated`, and `layoutMarkupBalloons()` reports it. No page is added.

**Leaders and marks.**

- Leaders are dotted hairlines (0.2pt) in the author's colour. Each runs along the anchor's line to
  4.6pt past the text column, then to the balloon, meeting it 6pt below its top.
- A run change is marked by a 3pt triangle under its first character. A paragraph change is marked
  by one at the text column's edge.
- A commented range is tinted and bracketed in its author's colour.

**PageMap.** The PageMap shape is unchanged. Pages keep the paper size. Fragments are normalized
against the page box, so they record where the shrunken content prints. The markup area, balloons
and leaders are excluded (`data-page-map-exclude`), and so are the comment registry's sources, as
with `"outside"`. Comment citations are therefore unavailable under this placement.

### Export: passes the option through

- `ExportOptions.commentColumnPlacement` is validated, combinations included, in `normalizeOptions`
  and again in the npm-export Node preflight, so the CLI rejects a bad combination before it
  launches a browser.
- The host request and the CLI accept it.
- Balloons put the converter in the Word presentation for every review profile.
- The report and the layout digest name the option only when it is `"balloons"`, so default reports
  and digests are unchanged. `render-report-v2.schema.json` lists the key.
- **Fail-closed.** The comment sources are excluded from the PageMap under this placement, so the
  export checks balloon coverage itself. If a comment marked in the document got no balloon, the
  export fails with `output_verification_failure`.
- **Truncation warning.** If any balloon was cut short, the report carries a
  `comment_balloon_truncated` warning. Its `resource` lists the comment ids, and its message counts
  comments and formatting changes. Warning codes are open strings in the report schema, so the
  schema itself did not change.
- No C#, WASM, stdio-host, Python or MCP surface changes: `RevisionPresentation` already flows
  through all of them.

## Default output

- **Docxodus presentation, any placement other than `"balloons"`.** Converter HTML is byte-identical
  to `main`. A main-versus-branch conversion of the whole `TestFiles` corpus showed no differences.
  `MarkupBalloonTests.DocxodusPresentation_RecordsNoBalloonData_AndKeepsItsTooltips` guards it.
- **Word presentation without balloons.** The rendering differs only by the recalibrated palette and
  change bar. The balloon data attributes are inert:
  `markup-balloons.spec.ts` paginates the same HTML with and without them under the margin placement
  and compares the page trees.
- **Export.** A default export's report names no `commentColumnPlacement`
  (`export-markup-balloons.spec.ts`).

## Wording

These are the wordings the converter writes after `Formatted:`. Properties that changed are named
with their new values. Run properties come before paragraph properties, in the order below, joined by
commas.

**Runs.** These form one `Font:` group, then the remaining run properties:

| Property | Wording |
|---|---|
| Typeface | `Font: (Default) Arial`, or `(Asian) …` for the East Asian slot; theme fonts are named by their typeface |
| Size | `16 pt` |
| Toggles | `Bold` / `Not Bold`, `Italic`, `All caps`, `Small caps`, `Hidden` |
| Vertical position | `Superscript` / `Subscript` / `Not Superscript/ Subscript` |
| Colour | `Font color: Red`, from Word's standard-colour names; `Text 1`, `Accent 2`… for theme colours, with `, Darker 25%` or `, Lighter 40%` for a theme shade or tint; `Custom Color(RGB(r,g,b))` otherwise |
| Underline | `Underline` / `Double underline` / `No underline` |
| Strikethrough | `Strikethrough`, `Double strikethrough` (or `Not …`) |
| Highlight | `Highlight` / `Not Highlight` |

**Paragraphs.**

| Property | Wording |
|---|---|
| Style | its name, for example `heading 1` |
| Alignment | `Centered` / `Right` / `Justified` / `Distributed` / `Left` |
| Indent | `Indent: Left: 0.5", Hanging: 0.25"`. Left is where a hanging first line starts; `First line:` and `Right:` follow when non-zero |
| Spacing | `Space Before: 12 pt, After: 6 pt` |
| Line spacing | `Line spacing: single` / `1.5 lines` / `Double` / `Multiple 1.15 li` / `At least 14 pt` / `Exactly 18 pt` |
| Numbering | `Numbered + Level: 1 + Numbering Style: 1, 2, 3, … + Start at: 1 + Alignment: Left + Aligned at: 0.25" + Indent at: 0.5"`, `Bulleted + Level: 1 + Aligned at: … + Indent at: …`, or `No bullets or numbering` |
| Toggles | `Don't add space between paragraphs of the same style` (or `Add …`), `Keep with next`, `Keep lines together`, `Page break before` (or their negations) |

Indentation that comes with newly applied numbering, or with a new list level, is the numbering's
own and is not described. A move to another level of the same list is left undescribed.

*Pending*, not yet measured:

- character spacing;
- character-style naming;
- shading;
- borders;
- tab stops;
- widow control;
- the label of a comment with no `w:initials`;
- landscape pages;
- the change mark of a non-empty paragraph.


## Ripple

| Area | Files |
|---|---|
| Core | `WmlToHtmlConverter.cs`, `Internal/WordFormatWording.cs` |
| Tests | `MarkupBalloonTests.cs` |
| Paginator | `npm/src/pagination.ts` (placement value, `deferMarkupBalloons`, `layoutMarkupBalloons()`), `npm/src/markup-balloons.ts` |
| Export | `npm/src/export-browser.ts`, `npm-export/src/{index,host,cli}.ts`, `docs/schemas/render-report-v2.schema.json` |
| Specs | `markup-balloons.spec.ts`, `export-markup-balloons.spec.ts`; existing change-bar expectations moved to 30pt |
| Bundle budget | `pagination.bundle.js` is 27.6 KiB gzip; its budget rose from 25 to 29 KiB |
