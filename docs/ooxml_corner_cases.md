# OOXML Corner Cases

This document tracks edge cases and quirks in Open XML document processing where Word's behavior differs from a strict interpretation of the specification, or where the specification is ambiguous.

## Table of Contents

1. [Numbering and Lists](#numbering-and-lists)
   - [Legal Numbering with Multi-Level Format Strings](#legal-numbering-with-multi-level-format-strings)
   - [List Numbering under Tracked Changes](#list-numbering-under-tracked-changes-deleted-paragraphs-dont-consume-numbers)
   - [Unresolvable Numbering vs. Removed Numbering (style indentation)](#unresolvable-numbering-vs-removed-numbering-style-indentation)
   - [`w:numberingChange/@w:original` holds at most 15 characters](#wnumberingchangeworiginal-holds-at-most-15-characters)
2. [Footnotes](#footnotes)
   - [Footnote Count Discrepancy in Legal Templates](#footnote-count-discrepancy-in-legal-templates)
3. [Package Output](#package-output)
   - [Misleading Deflate Hints Cause Compression Loss](#misleading-deflate-hints-cause-compression-loss)
   - [Content cloned into another part loses its namespace declarations](#content-cloned-into-another-part-loses-its-namespace-declarations)
4. [Paragraph Layout](#paragraph-layout)
   - [`w:lineRule="auto"` is a multiple of the FONT's line box, not of font-size](#wlineruleauto-is-a-multiple-of-the-fonts-line-box-not-of-font-size)
   - [An accumulated line-spacing error can resemble a top-margin deviation](#an-accumulated-line-spacing-error-can-resemble-a-top-margin-deviation)
   - [Cached TOC field results suppress hyperlink presentation](#cached-toc-field-results-suppress-hyperlink-presentation)
   - [`w:ind` has two spellings for each edge: `w:start`/`w:left` and `w:end`/`w:right`](#wind-has-two-spellings-for-each-edge-wstartwleft-and-wendwright)
5. [Theme Colors](#theme-colors)
   - [`w:color`/`w:fill` are a CACHE; `w:themeColor`/`w:themeFill` are the authority](#wcolorwfill-are-a-cache-wthemecolorwthemefill-are-the-authority)
6. [Contributing](#contributing)

---

## Numbering and Lists

### Legal Numbering with Multi-Level Format Strings

**Status:** Fixed (December 2024)
**Discovered:** 2024-12-22
**Test File:** `NVCA-Model-COI-10-1-2025.docx`

#### The Problem

When a paragraph uses a deeper indentation level (`ilvl`) with a format string that references parent levels (e.g., `%1.%2`), Word may not display the parent level numbers as expected.

#### Example Document Structure

```xml
<!-- abstractNum 3, level 0 -->
<w:lvl w:ilvl="0">
  <w:start w:val="1"/>
  <w:numFmt w:val="decimal"/>
  <w:lvlText w:val="%1."/>
</w:lvl>

<!-- abstractNum 3, level 1 -->
<w:lvl w:ilvl="1">
  <w:start w:val="4"/>
  <w:isLgl/>
  <w:numFmt w:val="decimal"/>
  <w:lvlText w:val="%1.%2"/>
</w:lvl>
```

Document paragraphs:
```
Para 1: ilvl=0, numId=3  → Word displays: "1."
Para 2: ilvl=0, numId=3  → Word displays: "2."
Para 3: ilvl=0, numId=3  → Word displays: "3."
Para 4: ilvl=1, numId=3  → Word displays: "4." (NOT "3.4")
```

#### Expected vs Actual Behavior

| Renderer | Item 4 Output | Notes |
|----------|---------------|-------|
| Microsoft Word | `4.` | Period at end, not middle |
| LibreOffice Writer | `4.` | Matches Word |
| LibreOffice HTML export | `4` | Uses `<ol start="4">` |
| Docxodus (current) | `3.4` | Incorrect - includes parent level |

**Key observation**: Word outputs "4." (number then period) even though level 1's format string is `%1.%2` (which would produce "3.4" if evaluated literally).

Note that level 0's format string is `%1.` (number then period). This suggests Word may be:
1. Detecting that item 4 at `ilvl=1` is an "orphan" (no proper parent-child nesting)
2. Falling back to level 0's format `%1.` but using the level 1 counter (4)
3. Result: "4." - which matches the observed output!

This "orphan detection" hypothesis would explain the behavior: Word recognizes when a deeper-level item doesn't have proper hierarchical nesting and reverts to simpler formatting.

#### Analysis

Our converter (`ListItemRetriever.cs`) builds `levelNumbers` for each paragraph by:

1. For `ilvl=1`, looping from level 0 to level 1
2. For level 0: inheriting the counter from the previous paragraph (3)
3. For level 1: using the `start` value (4)
4. Result: `levelNumbers = [3, 4]`
5. Format `%1.%2` produces: `"3" + "." + "4"` = `"3.4"`

Word appears to use different logic where:
- The `%1` token in the format string is either:
  - Omitted when there's no "active" parent paragraph at that level
  - Or interpreted differently when transitioning level depths

#### Potential Causes

1. **Orphan nesting detection**: Word may detect that para 4 at `ilvl=1` doesn't have a proper parent-child relationship with para 3 at `ilvl=0` (they're effectively siblings in a flat list that happens to use different levels).

2. **Level entry tracking**: Word may only include `%N` tokens in the output when level N has been "entered" as part of the current nesting chain, not just referenced from previous items.

3. **Start value heuristics**: When a level's `start` value (4) suggests continuation of an overall sequence, Word may apply special formatting rules.

#### Relevant Code

- `Docxodus/ListItemRetriever.cs`:
  - `FormatListItem()` (lines 1100-1144): Processes `lvlText` format tokens
  - Level number calculation (lines 980-1079): Builds `levelNumbers` array

```csharp
// Current logic in FormatListItem:
int levelNumber = levelNumbers[indentationLevel];
// This always uses the levelNumbers array, even if the level wasn't "entered"
```

#### The Fix

**Implementation**: Added "continuation pattern" detection in `ListItemRetriever.cs`.

**Detection criteria**:
A paragraph at `ilvl > 0` is in a "continuation pattern" when:
1. It's the first paragraph at this level in the current sequence, AND
2. The level's `start` value equals the parent level's counter + 1 (continues the sequence)

OR it inherits continuation status from a previous paragraph at the same level.

**What the fix does**:
When a continuation pattern is detected, the converter uses level 0's properties instead of the declared level's:
- Format string (e.g., `%1.` instead of `%1.%2`)
- Run properties (e.g., no underline instead of underline)
- Paragraph properties (e.g., tab stops and indentation)

**Code changes**:

1. **`Docxodus/ListItemRetriever.cs`**:
   - Added `ContinuationInfo` annotation class to track continuation state per paragraph
   - Added `GetEffectiveLevel()` helper method that returns 0 for continuation patterns
   - In `InitializeListItemRetriever`, after calculating `levelNumbers`:
     ```csharp
     // Detection logic
     if (levelNumbers[ilvl] == startValue && startValue == levelNumbers[ilvl - 1] + 1)
     {
         isContinuation = true;
     }
     ```
   - In `RetrieveListItem`, uses level 0's format string with current level's counter

2. **`Docxodus/FormattingAssembler.cs`**:
   - `NormalizeListItemsTransform`: Uses `GetEffectiveLevel()` to get list item level's rPr
   - `ParaStyleParaPropsStack`: Uses `GetEffectiveLevel()` to yield correct level's pPr and rPr
   - `AnnotateParagraph`: Uses `GetEffectiveLevel()` for numbering paragraph properties

**Result**:
- Items that continue a flat list sequence now render correctly (e.g., "4." instead of "3.4")
- Formatting (underline, bold, etc.) from the effective level is applied consistently
- Tab stops and indentation match the effective level's paragraph properties

#### Test Cases Needed

1. Standard multi-level list with proper nesting (1., 1.1, 1.2, 2., 2.1)
2. "Orphan" nesting like the NVCA example (1., 2., 3., then jump to level 1)
3. Legal numbering (`isLgl`) vs non-legal numbering behavior
4. Various `start` values and how they affect parent level display

#### References

- [ECMA-376 Part 1, Section 17.9.10 - lvlText](https://www.ecma-international.org/publications-and-standards/standards/ecma-376/)
- [ECMA-376 Part 1, Section 17.9.9 - isLgl](https://www.ecma-international.org/publications-and-standards/standards/ecma-376/)

### List Numbering under Tracked Changes (deleted paragraphs don't consume numbers)

**Status:** Fixed (August 2026)
**Discovered:** 2026-08-02 (the README's NVCA voting-agreement marquee redline)
**Test:** `Docxodus.Tests/TrackedChangesNumberingTests.cs` (TCN001–TCN006)

#### The Problem

In a tracked-changes document, which paragraphs consume list numbers? A literal reading of
the numbering spec says nothing about revisions — every `w:p` with a resolvable `w:numPr`
is a list item. Rendering a comparer-produced redline that way numbers deleted, moved-away
and inserted paragraphs in one continuous sequence, so every number after a deletion
disagrees with the final document.

#### Minimal XML Reproducer

A numbered list `Alpha, Bravo, Charlie` where `Bravo` is fully deleted (content in
`w:del`, pilcrow marked deleted — the shape both `WmlComparer` and `DocxDiff` emit):

```xml
<w:p><!-- Alpha: plain list item --></w:p>
<w:p>
  <w:pPr>
    <w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>
    <w:rPr><w:del w:id="1" w:author="Reviewer" w:date="..."/></w:rPr>
  </w:pPr>
  <w:del w:id="2" w:author="Reviewer" w:date="...">
    <w:r><w:delText>Bravo</w:delText></w:r>
  </w:del>
</w:p>
<w:p><!-- Charlie: plain list item --></w:p>
```

#### Renderer Comparison

| Renderer | Alpha | Bravo (deleted) | Charlie |
|----------|-------|-----------------|---------|
| Microsoft Word (All Markup) | `1.` | `2.` (struck) | `2.` |
| LibreOffice Writer (Show Changes) | `1.` | `2.` (struck) | `2.` |
| Docxodus before fix | `1.` | `2.` (struck) | `3.` |
| Docxodus after fix | `1.` | `2.` (struck) | `2.` |

#### Analysis

Word ties numbering to the **paragraph mark**. A pilcrow marked deleted
(`w:pPr/w:rPr/w:del`, or `w:moveFrom` on a move source) merges its paragraph into the
successor when the revision is accepted, so the paragraph does not survive as a numbered
item — and Word renumbers the *display* as if every change were already accepted. The
deleted paragraph still shows the value the counter holds at its position, without
advancing it; the next live paragraph shows the same value. That is the famous
"duplicate numbers next to a struck paragraph" Word renders in All Markup view, and it
is why a lawyer reading a Word redline sees final-document numbering. Inserted pilcrows
(`w:ins`/`w:moveTo`) count normally — they are part of the final document.

A second trap lives in the *styling* of the number glyph. When the paragraph **before**
a list item carries an inserted pilcrow, that can mean two very different things:

- **A split**: the ins-marked mark was inserted into *pre-existing* content (Enter
  pressed mid-paragraph with track changes on). The FOLLOWING paragraph occupies the
  newly created list position — its number is the "new" one.
- **A wholly inserted paragraph** (comparer output for a new or moved-in paragraph):
  own ins pilcrow AND all content inserted. The insertion is self-contained; rejecting
  it leaves the following paragraph untouched, so its number must NOT be styled as an
  insertion (or attributed to the revision's author).

`FormattingAssembler`'s marker-styling heuristic treated both alike, painting the
unchanged paragraph after every comparer-inserted paragraph as if its number were newly
inserted.

A comparison has one additional fact that an arbitrary tracked-changes document does not:
the resolved marker on both the original and revised paragraph. `DocxDiff` now preserves a
changed original marker in native `w:numPr/w:numberingChange[@w:original]` metadata. The HTML
tracked-changes renderer consumes it as a deleted old marker followed by an inserted current
marker, so a cascade such as `(a)` → `(b)` is visible instead of silently displaying only the
final `(b)`. For a wholly deleted or moved-from list paragraph, the same metadata keeps the
source-side marker visible rather than showing a counter value recomputed in the merged redline.

#### Relevant Code

- `ListItemRetriever.InitializeListItemRetrieverForStory` — the counting loop;
  `ParagraphMarkIsDeleted` paragraphs get their `LevelNumbers` annotation (they still
  render a struck number) but restore all forward-carried state (`previous`,
  start-override consumption, continuation tracking).
- `FormattingAssembler.NormalizeListItemsTransform` — the previous-paragraph-ins
  heuristic now requires the predecessor to carry pre-existing content
  (`IsWhollyInsertedParagraph`), and renders `w:numberingChange` as an old/new marker pair.
- `IrMarkupRenderer.StampResolvedNumberingChange` — compares the IR reader's already-resolved
  left/right markers and emits `w:numberingChange` when they differ.

The behavior is unconditional (no setting): the default HTML render path accepts
revisions before numbering runs, so only tracked-changes renders (and live
tracked-changes editing sessions) can observe it, and as-if-accepted is Word's reading
of the format.

### Unresolvable Numbering vs. Removed Numbering (style indentation)

**Status:** Fixed (September 2026, issue #821)

A paragraph whose `w:numId` is `0` has its numbering removed, and the indentation its
paragraph style supplies goes with it. The question is what happens when the numbering
reference is merely broken rather than explicitly removed.

#### Minimal XML reproducer

```xml
<!-- styles.xml -->
<w:style w:type="paragraph" w:styleId="Ind"><w:pPr><w:ind w:left="2880"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="IndDangling">
  <w:pPr><w:numPr><w:numId w:val="99"/></w:numPr><w:ind w:left="2880"/></w:pPr>
</w:style>
<!-- numbering.xml: numId 1 is a valid list; numId 2 names an abstractNum that doesn't exist -->
<w:num w:numId="2"><w:abstractNumId w:val="7"/></w:num>
<!-- document.xml -->
<w:p><w:pPr><w:pStyle w:val="Ind"/><w:numPr><w:numId w:val="0"/></w:numPr></w:pPr>...</w:p>
<w:p><w:pPr><w:pStyle w:val="Ind"/><w:numPr><w:numId w:val="99"/></w:numPr></w:pPr>...</w:p>
<w:p><w:pPr><w:pStyle w:val="Ind"/><w:numPr><w:numId w:val="2"/></w:numPr></w:pPr>...</w:p>
<w:p><w:pPr><w:pStyle w:val="Ind"/><w:numPr><w:ilvl w:val="10"/><w:numId w:val="1"/></w:numPr></w:pPr>...</w:p>
<w:p><w:pPr><w:pStyle w:val="IndDangling"/></w:pPr>...</w:p>
```

#### Comparison (left indent; the style says 2in)

| Paragraph numbering | LibreOffice 25.8 | Docxodus before | Docxodus after |
|---------------------|------------------|-----------------|----------------|
| `numId 0` | 0 | 0 | 0 |
| `numId 99` (no `w:num`) | 0 | 0 | 0 |
| `numId 2` (`w:num` without its `w:abstractNum`) | numbered "1.", 0.5in | 0 | 2in |
| `ilvl 10` (out of range) | numbered, 0.5in | 0 | 2in |
| style's own `numId 99`, paragraph has no `numPr` | 2in | 0 | 2in |

Word was not available to check. LibreOffice's DOCX import is written to match Word, and its
rule is explicit in `writerfilter/dmapper/DomainMapper.cxx` (`LN_CT_NumPr_numId`): when a
paragraph's `numId` names no list ("eg. disabled numbering using non-existent numId 0"), it
zeroes the inherited left and first-line indentation. `0` is just the one numId guaranteed not
to exist. The rule applies only to a paragraph's own `numPr`, not to a style's. When the `w:num`
exists but its definition is broken, LibreOffice treats the paragraph as numbered and invents
default levels; Docxodus can't render numbering it can't resolve, so it keeps the style's
indentation, which is the part of that behaviour that is well defined.

#### Relevant Code

- `ListItemRetriever.NotAListItem` (sets `IsZeroNumId`) vs. `ListItemRetriever.MalformedNumbering`
  (doesn't). `InitializeParagraphListItemSource` picks between them by whether the numId names
  a `w:num`; a style's unresolvable numbering and a level with no `w:lvl` always get
  `MalformedNumbering`.
- `FormattingAssembler` removes the rolled-up style `w:ind` only when `IsZeroNumId` is set.

#### Open question

A style whose own `numPr` says `numId 0` still has its indentation removed. LibreOffice keeps a
style's own `w:ind` in that case but drops one inherited through `w:basedOn`. That is well-formed
numbering, not a malformed reference, and is left as is.

### `w:numberingChange/@w:original` holds at most 15 characters

**Status:** Fixed (lossy by necessity)<br>
**Issue:** #861<br>
**Tests:** `Docxodus.Tests/DocxDiffNumberingChangeOriginalTests.cs`

#### The problem

When an insertion, deletion or move shifts a list's counter, a comparison records the label each
affected item displayed before the change in `w:numberingChange/@w:original`, so a redline can
show the old number struck through beside the new one. The schema types `w:original` as a string
of at most 15 characters. Labels routinely run longer: a spelled-out number (`cardinalText`,
`ordinalText`), a letter format far into a list (`lowerLetter` repeats the letter, so item 390 is
`zzzzzzzzzzzzzzz.`), or any level whose `w:lvlText` is itself long:

```xml
<w:lvl w:ilvl="0"><w:start w:val="3500"/><w:numFmt w:val="cardinalText"/><w:lvlText w:val="%1"/></w:lvl>
...
<!-- before the fix: an item was inserted above the list's first item -->
<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/>
  <w:numberingChange w:id="3" w:author="..." w:date="..." w:original="Three thousand five hundred"/>
</w:numPr>
```

#### What the value holds

ECMA-376 Part 1 §17.13.5.28 is cited in issue #861 as describing `w:original` as the original
number, with `%1`-style placeholders standing in for level values (the spec text was not
re-checked for this entry). Word-authored files do not follow that shape: the only
Word-written examples in `TestFiles` (`RP/RP026-NumberingChange.docx`) store the displayed label
itself (`w:original="1)"`, `"2)"`). Docxodus's own reader of the value, the tracked-change list
marker in `FormattingAssembler`, also treats it as the literal old label and renders it as the
deleted marker. A placeholder encoding would therefore render as `%1` in our own HTML and would
not match what Word writes, and it cannot shorten a long literal `w:lvlText` anyway.

| Consumer | Label longer than 15 characters |
|---|---|
| Open XML SDK 3.5 validator (Office 2019) | `Sem_AttributeValueDataTypeDetailed` on `w:numberingChange` before the fix; no error after |
| Word | not verified |
| LibreOffice | not verified |
| Docxodus before the fix | stored the full label (schema-invalid) |
| Docxodus after the fix | stores the first 14 characters and `…` |

#### The fix

`IrMarkupRenderer.StampOriginalNumberingMarker` is the only writer of `w:numberingChange` in the
comparison output, and both the two-way renderer and the consolidate renderer
(`IrCompositeMarkupRenderer`, which dispatches to the same per-op emitters) reach it for aligned,
deleted and moved-away list items. It passes the label through
`IrMarkupRenderer.NumberingChangeOriginal`: a label of 15 characters or fewer is stored unchanged;
a longer one keeps its first 14 UTF-16 code units followed by `…` (U+2026), dropping a surrogate
pair that would straddle the cut. The beginning of a label is kept because that is the part a
reader recognises (`Three thousand…`), and the ellipsis marks the stored value as shortened rather
than presenting a wrong label as exact.

**The loss.** The rest of a long label is not recoverable from the output. In the redline HTML the
old marker of such an item reads `Three thousand…` rather than `Three thousand five hundred`, and a
wholly deleted or moved-away item, whose marker is taken from `w:original`, shows the shortened
label too. Current markers, the live numbering and every label of 15 characters or fewer are
unaffected.

---

## Footnotes

### Footnote Numbering Uses Raw XML IDs Instead of Sequential Display Numbers

**Status:** Fixed (December 2024)
**Discovered:** 2024-12-23
**Test File:** `Model-COI-10-24-2024.docx` (NVCA model legal document)

#### The Problem

Docxodus was displaying footnote numbers using raw XML `w:id` attribute values instead of sequential display numbers. Per ECMA-376, the `w:id` is a reference identifier (linking `footnoteReference` to footnote definitions), not the display number. Display numbers should be calculated sequentially based on the order footnotes appear in the document.

**Example:**
- Document has 91 footnotes with XML IDs 2-92 (IDs 0, 1 are reserved for separator types)
- Word/LibreOffice display: 1, 2, 3, ..., 91 (sequential)
- Docxodus (before fix): 2, 3, 4, ..., 92 (raw XML IDs)

#### ECMA-376 Specification

The ECMA-376 specification clarifies how footnote numbering works:

1. **`w:id` is a reference identifier, NOT the display number**
   - The `w:id` attribute on `<w:footnoteReference>` links to the footnote definition in `footnotes.xml`
   - IDs 0 and 1 are reserved for `separator` and `continuationSeparator` types
   - Content footnotes typically start at ID 2

2. **Display number is determined by document order**
   - The first `<w:footnoteReference>` in document flow displays as "1"
   - The second displays as "2", and so on
   - This is independent of the `w:id` value

3. **`w:customMarkFollows` attribute**
   - When present, suppresses automatic numbering
   - Used for custom footnote marks (symbols, letters, etc.)

#### The Fix

**Implementation**: Added `FootnoteNumberingTracker` class in `WmlToHtmlConverter.cs`.

**How it works**:
1. Before conversion, scan the document for all `footnoteReference` and `endnoteReference` elements in document order
2. Build a mapping from XML ID to sequential display number (1, 2, 3...)
3. Store the mapping as an annotation on the root element
4. Use the mapping when rendering footnote references (superscripts) and footnote list items

**Code changes in `Docxodus/WmlToHtmlConverter.cs`**:
- Added `FootnoteNumberingTracker` class (lines 655-681)
- Added `BuildFootnoteNumberingTracker()` method to scan document and build mapping
- Added `GetFootnoteNumberingTracker()` helper method
- Updated `ProcessFootnoteReference()` to use display numbers instead of XML IDs
- Updated `ProcessEndnoteReference()` similarly
- Updated `RenderFootnotesSection()` to order footnotes by document order and use display numbers
- Updated `RenderEndnotesSection()` similarly
- Updated `RenderPaginatedFootnoteRegistry()` for pagination mode

**Result**: Footnotes now display with correct sequential numbers (1, 2, 3...) matching Word/LibreOffice behavior.

#### References

- [ECMA-376 Part 1, Section 17.11.7 - footnoteReference](https://www.ecma-international.org/publications-and-standards/standards/ecma-376/)
- [ECMA-376 Part 1, Section 17.11.10 - footnotes part](https://www.ecma-international.org/publications-and-standards/standards/ecma-376/)

---

### `continuationNotice` Reserved Footnote Rides at a POSITIVE `w:id` (NVCA contract)

**Status:** Fixed (2026-06)
**Discovered:** 2026-06-23
**Test:** `DocxDiffFootnoteRobustnessTests.ReservedContinuationNoticeAtPositiveId_DoesNotCollideWithRenumberedRealFootnote`

#### The corner case

The reserved boilerplate footnotes are commonly assumed to occupy non-positive ids (`separator` = -1, `continuationSeparator` = 0), with content footnotes starting at id 2. But Word emits a **third** reserved note — `continuationNotice` — and the real NVCA model contract carries it at a **positive** id:

```xml
<w:footnote w:type="separator" w:id="-1">…</w:footnote>
<w:footnote w:type="continuationSeparator" w:id="0">…</w:footnote>
<w:footnote w:type="continuationNotice" w:id="1"><w:p/></w:footnote>   <!-- positive id! -->
<w:footnote w:id="2">…first content footnote…</w:footnote>
```

Any code that (a) treats a typed note as reserved/kept-verbatim and (b) renumbers *content* notes from 1 will re-mint id 1 for the first content note → a **duplicate `w:id`** colliding with `continuationNotice`. In `DocxDiff` this corrupted **every** edit of the contract (even body/format-only edits that never touch a footnote), because the post-render renumber pass (`IrMarkupRenderer.RenumberNoteIds`) walks body references and re-sequences ids in reference order.

#### Renderer comparison

| | Renders the duplicate? |
|---|---|
| Word | N/A (Word never produces the collision; it keeps content ids ≥ 2 disjoint from reserved) |
| LibreOffice | Silently drops/repairs the colliding definition on load (loss) |
| Docxodus (before fix) | Emitted two `<w:footnote w:id="1">` — schema-invalid (`Sem_UniqueAttributeValue`) |

#### The fix

`RenumberNoteIds` now starts the content-note counter **above the highest positive reserved id** (so `{-1, 0}`-only documents are unchanged, but a `continuationNotice` at 1 pushes content notes to start at 2). The renumbered range stays disjoint from the kept boilerplate ids. Relevant code: `Docxodus/Ir/Diff/IrMarkupRenderer.cs` (`RenumberNoteIds`, the `int next = …` seed).

### LibreOffice Re-Associates Footnote References to Definitions POSITIONALLY (orphaned-definition fidelity)

**Status:** Documented behavior (not a Docxodus defect)
**Discovered:** 2026-06-23 (headless-LibreOffice footnote backstop, `tools/diffharness/lo/lo_footnote_check.py`)

#### The corner case

When a document contains an **orphaned footnote definition** (a `w:footnote` whose `w:id` is no longer named by any body `w:footnoteReference` — e.g. after a paragraph carrying the reference is deleted/rewritten, leaving the definition behind), LibreOffice on import does **not** resolve the surviving reference to its definition by `w:id`. It re-associates references to definitions **positionally** (the *n*-th reference → the *n*-th definition), so it displays the *first* definition's text for the surviving reference and drops the trailing one.

This means a body that references footnote id `2` ("See Section 1.2…") with an orphaned id `1` ("Include this provision…") still present renders in LibreOffice as "Include this provision…" — the orphaned definition's text. The OOXML is fully schema-valid (unique ids, the surviving reference resolves to exactly one definition by id); Word honors the id. It is purely a LibreOffice import behavior.

#### Why this is NOT a `DocxDiff` corruption

`DocxDiff.Compare` faithfully reproduces the **right** document's footnote structure on `accept` (and the left's on `reject`). The orphaned definition is a property of the user's edited (right) document itself — the fixture/edit removed only the body *reference*, not the *definition*. Loading the `right` document and the `accept(Compare(left,right))` document in LibreOffice yields **identical** footnote rendering (same count, same text, same positional association), confirming `accept ≡ right` cross-renderer. The "wrong" text is LibreOffice's handling of that valid OOXML shape, applied equally to the target and to the diff's accept output. No loss, no repair, no divergence introduced by the engine.

#### Relevant code / verification

- `tools/diffharness/lo/lo_footnote_check.py` — headless-LibreOffice load + footnote-count/text report (the independent validity backstop).
- `DocxDiffScenarioTests.Scenario_PreservesFootnoteStructure` — the in-process id↔reference↔text round-trip oracle (asserts at the OOXML id level, immune to LibreOffice's positional quirk).

### Word's Accept of a Tracked Note Deletion Leaves a CHILDLESS `<w:footnote/>` Shell

**Status:** Documented behavior; Docxodus deliberately differs (issue #631)
**Discovered:** 2026-08-31 (while making the stateless accept reproduce the counterpart's note store)
**Test File:** `TestFiles/RP/RP050-Deleted-Footnote.docx` + its `-Accepted.docx` oracle

#### The corner case

Word represents a tracked footnote deletion by marking the citation **and the whole definition**.
Minimal reproducer (the shape `RP050-Deleted-Footnote.docx` carries, trimmed):

```xml
<!-- word/document.xml: the citation is deleted -->
<w:del w:author="Eric White"><w:r><w:footnoteReference w:id="1"/></w:r></w:del>

<!-- word/footnotes.xml: every run is deleted AND the paragraph mark is deleted -->
<w:footnote w:id="1">
  <w:p>
    <w:pPr><w:rPr><w:del w:author="Eric White"/></w:rPr></w:pPr>
    <w:del w:author="Eric White">
      <w:r><w:footnoteRef/></w:r>
      <w:r><w:delText>This is a test.</w:delText></w:r>
    </w:del>
  </w:p>
</w:footnote>
```

Accepting that deletion, per renderer:

| | Accepted `word/footnotes.xml` |
|---|---|
| Word (the `RP050-…-Accepted.docx` oracle) | separators + a **childless `<w:footnote w:id="1"/>` shell** |
| LibreOffice | removes the definition; a leftover shell would shift its positional reference→definition pairing (entry above) |
| Docxodus (since #631) | separators only — the definition is **removed** |
| Docxodus (before #631) | separators + the childless shell, matching Word byte-shape but shipping debris |

`CT_FtnEdn` requires at least one block-level child, so the shell Word writes is
schema-questionable, invisible in Word (nothing references it), and — per the
positional-re-association entry above — actively dangerous in LibreOffice.

#### What Docxodus does

`RevisionProcessor.AcceptRevisions` removes a definition that the accept left both **orphaned**
(cited before, uncited after) and **blockless** (which can only happen when every block was
deletion-marked — an unmarked paragraph always survives). That is what `DocxSession`'s resolve
paths have done since #516, it is the schema-valid form, and it makes `Accept(Compare(l, r)) ≡ r`
hold for the counterpart's note store. The reversibility sweep (`RRS001`) stays green across the
divergence — RP050 is pinned `acceptEquivalent: false` there, so the sweep asserts recovery
through the coarser story-text check rather than modeled-semantic equivalence, and neither a
childless shell nor its absence contributes story text.

#### The cited-but-stripped shape is preserved, not repaired

A definition can end up bare while its citation survives: hand-built markup (content wholly
deletion-marked, citation kept), an ordinary partial resolution (record a note, accept only the
citation's revision, then run the stateless reject over what remains), or simply the input — the
corpus genuinely contains cited childless notes (`WC064-Footnote` ships one, and
`WC063-Footnote-Mod`'s note is childless). A cited `<w:footnote/>` with no block child is
schema-degenerate under `CT_FtnEdn`, but faithful reproduction of exactly that shape is what
`Accept(Compare(l, r)) ≡ r` means for those fixtures: inserting a "repair" paragraph was tried
during #636's review round and broke the WC063/WC064 round-trips in both directions. The
orphan-scoped prune never touches a cited note, so the shape flows through resolution unchanged;
whether Word repairs it on open is Word's business, and the engine will not manufacture content
the counterpart does not have.

#### Relevant code

- `Docxodus/Internal/NoteReferenceOps.cs` (`PruneNotesEmptiedByResolution`, the blockless guard)
- `Docxodus/RevisionProcessor.cs` (accept-side capture/prune)
- `DS430`/`DS432`/`DS434` in `DocxSessionRevisionTests.cs` — the husk shapes and their fates

### `OpenXmlValidator` Does NOT Resolve Note-Body (note-in-note) References — a validation blind spot

**Status:** Documented gotcha
**Discovered:** 2026-06-23 (non-body scope fidelity audit)

#### The corner case

A footnote/endnote definition body may itself contain a `w:footnoteReference`/`w:endnoteReference` (a note that cites another note — "note-in-note"). The SDK `OpenXmlValidator` validates references in the **document body** against the notes part, but does **not** resolve references that live **inside a note definition body**. So a *dangling* nested reference (one pointing to a note id that no longer exists after renumbering) produces **zero** schema errors — the validator simply does not check it.

This is a trap for any pipeline that uses "no new `OpenXmlValidator` errors" as its footnote-integrity oracle: it will pass a document whose note-in-note references dangle. In Docxodus this masked a real `DocxDiff` bug where `RenumberNoteIds` renumbered a body-referenced note's definition (e.g. id 5 → 2) but left a nested reference to it (inside another note's body) at the stale id 5.

#### How to actually catch it

Resolve **every** `footnoteReference`/`endnoteReference` in the document — body **and** inside every note definition body — against the note part's definition ids yourself; do not rely on the validator. See `DocxDiffFootnoteRobustnessTests.AllUnresolvedFootnoteRefs` (counts unresolved references across both scopes) and the fix in `IrMarkupRenderer.RenumberNoteIds` (records each definition's old→new id and remaps nested references).

#### Wrinkle (2026-06-24): it also FALSE-POSITIVES, and the value it names follows a renumber

The blind spot is worse than "ignores them": `OpenXmlValidator` (Office2019) emits a `Sem_MissingReferenceElement` for a note-in-note reference **even when the target definition is present** (a false positive — observed on `TestFiles/DD/DD001-DenseBookmarkXrefFootnote.docx`, whose footnote 2 cites footnote 5, where 5 exists). The error's `Description` embeds the *reference value* (`…The reference value is '5'.`). So when `DocxDiff` **correctly** compacts a gapped note id (5 → 4), the validator's false positive simply re-emits with the new value (`'4'`), at the same part/path.

This is a trap for a schema-error oracle that diffs validator output across input↔output keyed on the description: the input copy (`'5'`) and output copy (`'4'`) look like *different* errors, so the legitimate renumber is mis-counted as a NEW defect. `DocxDiffBookmarkRealDocTests.SchemaErrors` defends against this by keying on `{Id}@{Part.Uri}` + a *value-normalized* description (`'\d+'` → `'#'`); genuine new dangling references in a different part are still surfaced, and real note-in-note resolution is checked structurally by the `UnresolvedNoteRefs` oracle (which does not consult the validator at all).

---

## Comments

### Comment threading is keyed on `w14:paraId`, NOT the comment `w:id` (and a dedup clone must carry its own paraId)

**Status:** Documented behavior + design note (comment fidelity campaign)
**Discovered:** 2026-06-24 (`DocxDiffCommentStructureTests`, headless-LibreOffice comment oracle `tools/diffharness/lo/lo_comment_check.py`)

#### The corner case

A threaded comment reply is linked to its parent **not** by the comment's `w:id`, but by the `w14:paraId` of the comment-definition paragraph: `commentsExtended.xml` carries `<w15:commentEx w15:paraId="…" w15:paraIdParent="…">` where both values are `w14:paraId`s of `<w:comment>/<w:p>` elements in `comments.xml`. Both Word and LibreOffice resolve "which comment is a reply to which" purely through this paraId graph. So renumbering a comment's `w:id` (as the `DocxDiff` dedup does for the del/ins copies of a rewritten commented paragraph — the comment analogue of the bookmark renumber-collision) does **not** by itself break threading.

The trap is in the **reverse** direction. When `DocxDiff` clones a comment definition to give the deleted (reject-side) copy a fresh `w:id`, a naive clone either (a) **duplicates** the original's `w14:paraId` — two comments now claim the same threading key — or (b) **strips** the paraId to avoid that duplicate, which silently severs the clone from `commentsExtended` so a **reject-side threaded reply dangles** (its `paraIdParent` no longer names a comment with that paraId). Both are wrong: (a) is ambiguous, (b) loses the reply→parent link on reject.

#### The fix

`IrMarkupRenderer.NormalizeComments` (phase B) gives each dedup clone a **fresh** `w14:paraId` (allocated above the max existing paraId) *and* clones the matching `commentsExtended`/`commentsIds` entry under the fresh paraId (`CloneThreadingEntryForParaId`), preserving `paraIdParent`. So the reject-side clone keeps its own threading link, exactly as the accept-side original keeps the unchanged one. Verified independently: `lo_comment_check.py` enumerates LibreOffice `Annotation` fields and asserts every reply's `ParentName` names a loaded comment — the dense fixture's Compare output reports 2 threaded replies (original + clone), both resolving.

#### `OpenXmlValidator` does NOT flag a duplicate `w14:paraId` (a second comment-threading blind spot)

Like the note-in-note blind spot above, the SDK `OpenXmlValidator` (Office2019) does **not** validate `w14:paraId` uniqueness across comment definitions — a document with two `<w:comment>/<w:p>` sharing one paraId is "schema-valid" to the validator but ambiguous to Word's/LibreOffice's threading resolver. So "no new validator errors" is **not** sufficient to prove comment-threading integrity; assert paraId/threading structurally (`DocxDiffCommentStructureTests.AnchorProjection` resolves each reply's parent through the paraId graph and checks `accept ≡ right` / `reject ≡ left` on the resolved-parent text).

#### v1 limitation: cross-document comment id / paraId collision (independent documents only)

The comment merge (`MergeRightCommentDefinitions`) and collapse assume a comment present in both sides carries the SAME `w:id` — true when the two inputs are two versions of ONE document (Word never reassigns a comment's id, so an edited doc's comment ids are stable). Two **independent** documents that each happened to assign `w:id="0"` to a DIFFERENT comment anchored on the same text, or a right-added comment whose `w14:paraId` GUID collides with a left comment's, are out of v1 scope: the cross-document case can leave `accept` showing the left comment's text (the right definition is not re-id'd and merged) or duplicate a `w14:paraId`. The output stays schema-valid and every reference resolves to exactly ONE comment (the `(C)` backstop guarantees that) — it is a content/threading-attribution gap, not a structural corruption, and it does not arise from the diff/review workflow the engine targets (before/after of one document). Re-id'ing right-sourced markers for a genuinely independent-document merge is a follow-on.

#### Unchanged comment ⇒ single BARE range (mirrors bookmarks)

A comment present in both sources whose anchored text is **un**edited collapses (phase A) to a single bare `commentRangeStart`/`End`/`commentReference` (no `w:ins`/`w:del` wrapper) so it survives **both** accept and reject — the same identity-aware collapse `NormalizeBookmarks` does. A right-**added** comment's bare markers (which landed in equal content) are instead wrapped in `w:ins` (phase A2) so the comment toggles with its side and does not leak into the reject (`reject ≡ left`); a left-**deleted** comment's are wrapped in `w:del`. LibreOffice drops a `commentReference` whose `w:comment` definition is missing (its own dangling-comment signal), so the oracle's clean load + refresh-stable comment count is the cross-renderer confirmation that every reference resolves.

---

## Table/Cell Width as Percent-Suffixed String (`w:tblW` / `w:tcW` with `w:type="pct"`)

**Status:** Fixed (2026-05) — Issue #210

### Symptom

`WmlToHtmlConverter.ConvertToHtml` (`convertDocxToHtml` in the npm wrapper) threw
`FormatException` — `Conversion failed: Format_InvalidStringWithValue, 100%` —
for any document whose table or cell width was a percentage.

### Minimal XML reproducer

```xml
<w:tbl>
  <w:tblPr>
    <!-- percent-suffixed string form -->
    <w:tblW w:w="100%" w:type="pct"/>
  </w:tblPr>
  <w:tr>
    <w:tc>
      <w:tcPr><w:tcW w:w="50%" w:type="pct"/></w:tcPr>
      <w:p><w:r><w:t>Item</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
</w:tbl>
```

### The corner case

The `w:w` attribute on `w:tblW` / `w:tcW` has schema type `ST_TblWidth`
(a union over `ST_MeasurementOrPercent` + `ST_DecimalNumber`). Under
`w:type="pct"` the value may be expressed **two** schema-valid ways:

| Form | Example | Meaning |
|------|---------|---------|
| Integer (fiftieths of a percent) | `w:w="5000"` | 5000 / 50 = 100% |
| Percent-suffixed string | `w:w="100%"` | a literal 100% |

Microsoft Word writes the integer-fiftieths form. The widely used `docx`
JavaScript library writes the **percent-suffixed string** form for
`WidthType.PERCENTAGE` — both are schema-valid, but Docxodus only handled the
integer form, casting the attribute straight to `int`. `(int)"100%"` throws.

### Renderer comparison

| Width markup | Word | LibreOffice | Docxodus (before) | Docxodus (after) |
|--------------|------|-------------|-------------------|------------------|
| `w:w="5000" w:type="pct"` | 100% | 100% | `width: 100%` | `width: 100%` |
| `w:w="100%" w:type="pct"` | 100% | 100% | **throws** | `width: 100%` |
| `w:w="9000" w:type="dxa"` | 450pt | 450pt | `width: 450pt` | `width: 450pt` |

### Relevant code

`Docxodus/WmlToHtmlConverter.cs` — `ParseTblWidthValue(XAttribute, out bool isExplicitPercent)`
centralizes the parse and is called from `ProcessTable` (table-level `w:tblW`)
and the cell-processing path (`w:tcW`). When the raw value ends with `%`,
`isExplicitPercent` is set and the number is treated as a literal percentage;
otherwise a `pct` value is divided by 50 (fiftieths -> percent) as before.
Non-numeric values return `null` and are skipped instead of throwing.

### Tests

`Docxodus.Tests/HtmlConverterTablePercentageWidthTests.cs`
(`HcTablePercentageWidthTests`).

---

## DocxDiff: zero-width markers that are NOT diff tokens (bookmarks, field plumbing, soft hyphens)

### Symptom

Diffing two DOCX with `DocxDiff` and editing a paragraph that carries a bookmark, a `REF`/`PAGEREF` field,
or a `w:noBreakHyphen`/`w:softHyphen`/`w:sym` produced output where, after accept or reject:

- a `w:bookmarkStart`/`w:bookmarkEnd` was **dropped** (orphaning the bookmark, dangling every
  `w:hyperlink @w:anchor` and `REF`/`PAGEREF`/`NOTEREF`/`HYPERLINK \l` reference that targets it),
- the same bookmark **id was duplicated** across the `w:del` and `w:ins` copy (`Sem_UniqueAttributeValue`),
- a whole `REF` **field vanished** when the text *before* it was edited, and
- a body character was **dropped** next to a non-breaking/soft hyphen (the reject of a
  "Company‑Controlled Intellectual" run lost the "I").

The SDK `OpenXmlValidator` caught only the duplicate id; the dropped marker / dropped field / dropped char are
schema-valid (the validator does not resolve cross-references), so they require a STRUCTURAL round-trip oracle
(bookmark id↔name↔reference integrity + body-text `reject ≡ left` / `accept ≡ right`).

### The corner case

The IR diff engine reconstructs an edited paragraph by **slicing the SOURCE run-level XML** at character
offsets the **token diff** decided. A run-level element is one of three kinds with respect to that offset
math:

| element | IR / tokenizer treats it as | source slicer must treat it as |
|---|---|---|
| `w:t` text | N chars | N chars (splittable) |
| `w:bookmarkStart`/`End`, `w:fldChar`, `w:instrText` | **dropped / 0 chars, NOT a token** | **0 chars but ALWAYS emitted** |
| `w:noBreakHyphen`/`w:softHyphen`/`w:sym` | **1 char of text** (e.g. U+2011) | **1 char** |

Two distinct bugs followed from violating that table:

1. **Boundary drop.** Because a bookmark/field marker is *not a diff token*, the token-driven
   boundary-ownership flags (`includeStart/EndZeroWidth`) were blind to it, so a marker sitting exactly at an
   edit boundary was claimed by neither adjacent op and disappeared. Fix: flag these markers `AlwaysKeep` in the
   slicer (taken anywhere in `[start,end]`), then reconcile context in post-render passes (`NormalizeBookmarks`,
   `NormalizeFields`).
2. **Off-by-one.** `w:noBreakHyphen`/`w:softHyphen`/`w:sym` ARE one character in the IR (the reader emits an
   `IrTextRun`), so the tokenizer counts them — but the slicer counted them as zero-width. Every such element
   shifted the slice by one and dropped an adjacent character. Fix: the slicer advances the char counter by one
   for them, matching the IR.

A bookmark/field present in BOTH documents is *unchanged by the edit* (only the surrounding text moved), so its
correct representation is a single **bare** (untracked) pair that survives both accept and reject — NOT a
tracked `w:ins`/`w:del` copy. `NormalizeBookmarks`/`NormalizeFields` collapse to that; a wholly inserted/deleted
bookmark or field keeps its revision context. Bookmarks nested in opaque content (`m:oMath`, `w:drawing`) are
deliberately left untouched — they are part of that element's canonical content hash, so renumbering them would
break `reject ≡ left`.

### Word vs LibreOffice

No genuine Word-vs-LibreOffice *divergence* was found for bookmark/cross-reference handling: with the fixes,
both the bookmark structural round-trip and (per the `lo_bookmark_check.py` oracle design) LibreOffice's own
`GetReference` field resolution agree that every reference resolves. The one non-divergent quirk worth noting:
the strict ECMA-376 schema rejects `<w:w w:val="0">` (character scale 0), which the real NVCA COI source carries
65× and which Word writes and tolerates — the diff merely relocates those runs, so it is a *source* quirk, not a
diff defect (it appears identically when validating the input).

### Relevant code

- `Docxodus/Ir/Diff/IrMarkupRenderer.cs` — `SourceRunModel` (`AlwaysKeep`/`FieldPlumbingKeep`, the 1-char
  hyphen/sym segment), `NormalizeBookmarks`, `NormalizeFields`, `ExpandFieldForRevision`.
- `Docxodus/Ir/IrReader.cs` — `EmitRunChild` (N7/N8: `noBreakHyphen`/`softHyphen`/`sym` → 1-char `IrTextRun`).
- `Docxodus/WmlComparer.cs` — `AddNumberingChildInSchemaOrder` (numbering-merge child order).

### Tests

`Docxodus.Tests/DocxDiffBookmarkStructureTests.cs` + `DocxDiffBookmarkFixtures.cs` (synthetic corpus),
`DocxDiffBookmarkRealDocTests.cs` (real NVCA COI/SPA), and the `bkmk-struct` column + `lo/lo_bookmark_check.py`
oracle in `tools/diffharness`.

---

## DocxDiff: `PreAcceptInputRevisions` accept-all flattens prior authorship

**Status:** Documented (2026-06) — the `revisionsInInput` campaign.

### The corner case

This is not a Word-vs-spec divergence but a **lossy-by-design transformation** worth pinning, because "just
accept-all both sides, then diff" looks innocent and is not. When an input is itself a redline (carries
un-accepted `w:ins`/`w:del`/`w:moveFrom`), DocxDiff's default already diffs the **accepted view** (rule N13 —
`IrReader` runs `RevisionView.Accept` before building the IR), so the produced *body* carries only the new
diff's revisions. But the output package is cloned on the LEFT input and only the body (+ changed notes) is
rebuilt, so pre-existing revision markup in **carried-over parts** (headers/footers, unchanged
footnotes/endnotes, styles, comments) is passed through verbatim. The opt-in
`DocxDiffSettings.PreAcceptInputRevisions` eliminates that by accepting BOTH whole inputs first (cleaning the
body, headers/footers, notes, and styles — the parts `RevisionProcessor.AcceptRevisions` processes; a tracked
change inside a comment *definition* or a glossary/building-blocks entry is NOT touched, since
`RevisionProcessor.AcceptRevisions` does not process the comments part or the `GlossaryDocumentPart`) — but
accept-all itself has two honest costs that a caller must understand before enabling it.

### Minimal XML reproducer

An input whose header carries a prior reviewer's tracked insertion (identical on both diff sides, so a pure
carry-over, not a diff):

```xml
<!-- left.docx and right.docx both contain this header part -->
<w:hdr>
  <w:p>
    <w:r><w:t xml:space="preserve">Header </w:t></w:r>
    <w:ins w:id="99" w:author="OldReviewer" w:date="2020-01-01T00:00:00Z">
      <w:r><w:t>CONFIDENTIAL</w:t></w:r>
    </w:ins>
  </w:p>
</w:hdr>
```

### Behavior table

| Setting | Output header | `accept(result)` header | `reject(result)` header | Round-trip in header? |
|---|---|---|---|---|
| default (`PreAcceptInputRevisions = false`) | `<w:ins author="OldReviewer">CONFIDENTIAL</w:ins>` **(leaked verbatim)** | `Header CONFIDENTIAL` | `Header ` (**leaked ins rejected → text dropped**) | **No** — `reject ≠ accept-view(left)` |
| `PreAcceptInputRevisions = true` | `Header CONFIDENTIAL` (plain, accepted) | `Header CONFIDENTIAL` | `Header CONFIDENTIAL` | **Yes** |

Word and LibreOffice behave the same on the *outputs* — there is no renderer divergence here. The divergence is
between the default's leaked, non-round-tripping header and the flag's clean one.

### Analysis — the two honest costs of accept-all

Even with the flag fixing the leak/round-trip, accept-all is opinionated and lossy and must not be enabled
silently:

1. **It flattens pre-existing authorship and change boundaries.** Accepting collapses each input's own tracked
   changes into final text. `OldReviewer` (and *where* their edit began/ended) is gone from the result; the
   output's authorship reflects only the new diff. You cannot recover "who edited what" afterward.
2. **"Accept all" is itself a policy.** Leaving a change in tracked form is how a reviewer defers or rejects it;
   accept-all overrides that, materializing every insertion and dropping every deletion regardless of the prior
   reviewer's intent. If the inputs' in-flight revisions must be preserved or re-adjudicated, resolve them by an
   explicit policy first, then diff — do not reach for `PreAcceptInputRevisions`.

### The Word-Combine alternative — `PreserveInputRevisions` and the one-sided Reject All

Word's **Combine** does neither the default's leak nor the flag's flatten: it **preserves** the inputs'
pre-existing tracked revisions in its output verbatim (original author/date markup intact) while the
text diff is computed over the accepted view. Verified against Word-oracle outputs that were later identified
as Combine-shaped: an input with 176 revisions by another author keeps them in the result alongside the fresh
revisions. (Word's **Compare** treats them as accepted instead — the flatten above.)
`DocxDiffSettings.PreserveInputRevisions` reproduces this (equal blocks + whole-block inserts, in the body
and in footnote/endnote bodies, in v1; it WINS
over `PreAcceptInputRevisions` when both are set). The `DocxCompare` front door does **not** enable it: the
oracle batch this was decoded from turned out to be Word *Combine* output, and the front door models Compare,
so it only pre-accepts — see the next section.

The Word behavior worth pinning here: **Reject All on such an output does NOT restore the left document.**
Rejecting a preserved foreign `w:del` RESTORES its deleted text (text the left side never showed), and
rejecting a preserved foreign `w:ins` removes text the accepted view carried. Word's Combine output behaves
identically under Reject All — the one-sided round trip (`accept ≡ right` holds, `reject ≠ left` where
foreign markup exists) is inherent to preserving input revisions, not a Docxodus defect. Do not "fix" it.

### Relevant code

- `Docxodus/DocxDiff.cs` — `DocxDiffSettings.PreAcceptInputRevisions` + the `PreAccept(...)` pre-pass wired into
  all seven entry points; `IrMarkupRenderer.Render` clones the output on the LEFT package (the carry-over source);
  `DocxDiffSettings.PreserveInputRevisions` (the Word-Combine opt-in, precedence over the pre-accept).
- `Docxodus/Ir/IrReader.cs` — `ApplyRevisionView` (rule N13: `RevisionView.Accept` before IR build).
- `Docxodus/Ir/Diff/IrMarkupRenderer.cs` — `BuildPreservedOriginalIndex` / `NormalizePreservedClone` + the
  preserve-aware `EmitVerbatim`/`EmitWholeBlock`/`MarkWholeParagraph`/`MarkParagraphMark`/`MarkWholeTable`.
- `Docxodus/DocxDiffCompatibility.cs` — the `revisionsInInput` catalog entry (now `Covered`).

### Tests

`Docxodus.Tests/Ir/Diff/RevisionsInInputDefaultTests.cs` (pins the default: clean body + leaking carry-over +
the broken header round-trip), `PreAcceptInputRevisionsTests.cs` (the flag is the wrapper, no stale
authorship, every-scope round-trip, schema validity, multi-author redline-of-a-redline), and
`DocxDiffPreserveInputRevisionsTests.cs` (preservation of foreign ins/del in equal + inserted blocks, no
same-kind nesting, the fully-deleted-paragraph ride-along, the pinned reject caveat, precedence).

---

## `DocxCompare`: the original document's own pending changes

**Status:** Documented (2026-09) — a deliberate difference from Word (issue #845).

### The corner case

When the **original** document already carries tracked changes by another author, Word's Compare keeps
that document's pending insertions marked as insertions in its result, so accept-all on Word's redline gives
the revised text *plus* those retained insertions. The `DocxCompare` front door (every transport's compare)
does not: it compares the **accepted view** of each document (`PreAcceptInputRevisions`), so an original's
pending change is resolved as accepted before the diff, nothing by the earlier author survives, and only the
differences between the two accepted views are marked, by the compare author.

### Minimal XML reproducer

```xml
<!-- original.docx: Alice's insertion is still pending -->
<w:p>
  <w:r><w:t xml:space="preserve">The quick </w:t></w:r>
  <w:ins w:id="1" w:author="Alice" w:date="2026-01-01T00:00:00Z"><w:r><w:t xml:space="preserve">brown </w:t></w:r></w:ins>
  <w:r><w:t>fox.</w:t></w:r>
</w:p>
<!-- revised.docx: <w:p><w:r><w:t>The quick fox.</w:t></w:r></w:p>  (Alice's insertion rejected) -->
```

### Behavior table (Docxodus, measured; `[+x]` inserted, `[-x]` deleted, both by the compare author)

| Original | Revised | Docxodus redline | accept all | reject all |
|---|---|---|---|---|
| `quick [+Alice: brown] fox` | `quick brown fox` (accepted) | `quick brown fox` | `quick brown fox` | `quick brown fox` |
| `quick [+Alice: brown] fox` | `quick fox` (rejected) | `quick [-brown] fox` | `quick fox` | `quick brown fox` |
| `quick [+Alice: brown] fox` | `quick brown fox jumps` | `quick brown fox[+ jumps]` | `… fox jumps` | `quick brown fox` |
| `quick [+Alice: brown] fox` | same pending insertion | `quick brown fox` | `quick brown fox` | `quick brown fox` |
| `The [-Alice: lazy] dog` | `The dog` (accepted) | `The dog` | `The dog` | `The dog` |
| `The [-Alice: lazy] dog` | `The lazy dog` (rejected) | `The [+lazy] dog` | `The lazy dog` | `The dog` |

Word, per the issue's observation of its compare output: the original's pending insertion `brown` stays a
tracked insertion in the redline, so in the "rejected" row Word's accept-all gives `quick brown fox`, not the
revised document's `quick fox`. (Docxodus cannot run Word; the Word column is the reported behavior, not a
measurement made here.)

### Analysis — why the difference is deliberate

Every Docxodus comparison surface promises one contract: **accept all ≡ the revised document, reject all ≡
the original** (each as accepted). The redline reversibility proof, the delivery bundle's change receipt and
the semantic change set all rest on it. Retaining the original's pending insertions breaks both halves in the
"rejected" rows: accept-all keeps text the revised document removed, and reject-all removes text both
documents' accepted views contain. A redline that cannot be reversed to either input is the wrong default for
a comparison API, so the front door keeps resolving input revisions as accepted. A caller who needs the inputs'
own revisions carried through can use the raw `DocxDiff` API with `PreserveInputRevisions` (which preserves
the *revised* document's pending changes; see the section above and its one-sided round trip). Reproducing
Word's exact retention as an opt-in profile is possible later if a caller needs Word's accept-all semantics.

### Relevant code

- `Docxodus/DocxCompare.cs` — `ApplyFrontDoorRevisionPolicy` (pre-accept only).
- `Docxodus/DocxDiff.cs` — `PreAcceptInputRevisions`, `PreserveInputRevisions`.

### Tests

`Docxodus.Tests/DocxCompareOriginalPendingRevisionsTests.cs` pins every row above through the front door
(redline markup, accept-all, reject-all, and that no change by the earlier author survives);
`DocxCompareTests.FrontDoorPolicy_*` pin the policy flags themselves.

---

## `*PrChange` inners are CT_*Base: rejecting a property change must NOT drop what lives outside it

**Discovered:** 2026-07-03, block-format-change family.

Word's block-level property-revision markers store the OLD properties in an inner element whose type is the `…Base` variant of the container — which **excludes** the sibling content that is not part of the tracked property change:

- `w:pPrChange`'s inner `w:pPr` is `CT_PPrBase` — no paragraph-mark `w:rPr`, and **no inline `w:sectPr`**.
- `w:sectPrChange`'s inner `w:sectPr` is `CT_SectPrBase` — **no `w:headerReference`/`w:footerReference`**.

### The trap

A naive reject implementation replaces the whole container with the inner:

```
// WRONG — drops the references / inline sectPr that were never part of the change
if (element.Name == W.sectPr && element.Element(W.sectPrChange) != null)
    return element.Element(W.sectPrChange).Element(W.sectPr);
```

Because the inner is reference-less, rejecting a `w:sectPrChange` deletes the section's headers and footers; rejecting a `w:pPrChange` on a section-final paragraph deletes the section break. This is a **silent** structural loss — the document still validates.

### Word's behavior

Word rejects the tracked *property* change (restores the old page setup / paragraph properties) while **keeping** the references and the inline `w:sectPr` intact — they were never revised.

| Reject a `w:sectPrChange` (section had a `w:headerReference`) | Header reference after reject |
|---|---|
| Word | preserved |
| Docxodus (before fix) | **dropped** |
| Docxodus (after fix) | preserved |

### The fix

`RevisionProcessor.RejectRevisionsForPartTransform` rebuilds the container: keep the CURRENT out-of-scope children (references for sectPr; the mark `w:rPr` + inline `w:sectPr` for pPr), restore the inner's in-scope properties.

### Relevant code

- `Docxodus/RevisionProcessor.cs` — the `w:sectPr`/`w:sectPrChange` and `w:pPr`/`w:pPrChange` reject branches.
- `Docxodus/Ir/Diff/IrMarkupRenderer.cs` — `ApplySectPrChange`/`ApplyBlockFormatChanges` produce the markers with reference-less CT_*Base inners (so they round-trip only *with* the fix).

### Tests

`Docxodus.Tests/Ir/Diff/BlockFormatChangeTests.cs` — `RejectRevisions_preserves_inline_sectPr_when_rejecting_a_pPrChange` and `SectPrChange_reject_preserves_header_footer_references`.

---

## Document metadata: `w:sectPr` in tables is a section break; `w:sectPr` in a text box is NOT

**Status:** Fixed (issue #51)

### Symptom

`WmlToHtmlConverter.GetDocumentMetadata()` reported the wrong number of sections
for documents whose section break lived inside a table cell — it only scanned the
body's direct children, so an in-cell `w:sectPr` was invisible and the section
count (and the per-section page dimensions / paragraph & table index ranges used
for lazy-loading pagination) were off.

### The corner case

A section break is expressed as a `w:sectPr` in the `w:pPr` of the paragraph that
*ends* the section (or the trailing `w:body/w:sectPr` for the final section).
While Word's UI does not let you place one inside a table, the OOXML is free to
carry a `w:sectPr` on a paragraph nested in a table cell, and a faithful metadata
scan must find it:

```xml
<w:body>
  <w:tbl>
    <w:tr><w:tc>
      <w:p><w:pPr>
        <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>   <!-- ends section 1 -->
      </w:pPr></w:p>
    </w:tc></w:tr>
  </w:tbl>
  <w:p><w:r><w:t>section 2</w:t></w:r></w:p>
  <w:sectPr><w:pgSz w:w="15840" w:h="12240"/></w:sectPr>          <!-- final section -->
</w:body>
```

The subtlety is what to do with a `w:sectPr` inside a **text box**. A text box
(`w:txbxContent`, reached through `w:r/w:pict/v:textbox` or the DrawingML
`wps:txbx`) is a **separate story**: it is floating content that does not
paginate the main document. A `w:sectPr` there is meaningless to main-document
pagination, and counting it would invent a phantom section and corrupt the page
dimensions reported for the surrounding content. The same is true of any story
reached only *through a run* — it must be excluded.

| Placement | Counts as a main-document section? |
|-----------|-----------------------------------|
| `w:body/w:p/w:pPr/w:sectPr` (body paragraph) | ✅ yes |
| `w:body/w:sectPr` (trailing) | ✅ yes |
| `w:tbl/w:tr/w:tc/w:p/w:pPr/w:sectPr` (table cell) | ✅ yes (this fix) |
| `…/w:txbxContent/w:p/w:pPr/w:sectPr` (text box) | ❌ no — separate story |
| header / footer / footnote / endnote / comment part | ❌ no — separate part, not in the body tree |

### The fix

`CollectSectionData` walks the body as a single **main story** in document order,
descending into tables (`w:tbl → w:tr → w:tc`) because a table cell's content is
part of that story, but treating every `w:p` as a **leaf** — only its direct
`w:pPr/w:sectPr` is inspected, never its runs. Because a text box is always nested
inside a run, this one rule excludes text-box section properties (and text-box
paragraphs) automatically, with no special-casing of `w:txbxContent`. Only
body-level tables count toward the table total (a nested table is part of its
containing cell's content), preserving the pre-fix counts for the common case.

Headers, footers, footnotes, endnotes and comments live in their own parts and
are never in the body tree, so they are out of scope by construction.

### Relevant code

`Docxodus/WmlToHtmlConverter.cs` — `CollectSectionData` (the recursive
`WalkMainStory` local function).

### Tests

`Docxodus.Tests/DocumentMetadataTests.cs` —
`DM022_GetDocumentMetadata_DetectsSectionBreakInsideTableCell` (in-cell break is
counted, table attributed to the section it starts in) and
`DM023_GetDocumentMetadata_IgnoresSectionPropertiesInsideTextBox` (a text-box
`w:sectPr` does not create a section, and text-box paragraphs are not counted).

---

## Tables: Word's compare backfills a hairline cell-margin/indent on fixed-width tables

**Status:** Fixed (DocxDiff — `WordCompareTableNormalizer`)

### Symptom

A DocxDiff redline of two documents containing a fixed-width table rendered (in LibreOffice) with every
cell's text shifted horizontally versus Microsoft Word's redline of the same pair — the whole table
"ghosts" against the oracle, costing large amounts of pixel-diff energy even when the tracked-changes
word stream is byte-identical to Word's.

### The corner case

A fixed-width table (`w:tblW w:type="dxa"`, cell widths summing to the table width) authored with **no**
explicit `w:tblCellMar` relies on the application's default cell margin. Word's and LibreOffice's defaults
disagree: LibreOffice insets cell text by ≈108 twips (its default), which on a fixed layout eats into the
declared column widths, while Word insets by a hairline.

Word's **compare** output does not leave the margin implicit — it materializes it. Across the Word-compare
corpus, every fixed-width table lacking cell margins came back with `w:tblCellMar` left/right **and** a
matching `w:tblInd` whose value equals the table's border width:

| Source `tblBorders` `w:sz` | Point size | Materialized `tblCellMar`/`tblInd` (twips) |
|---|---|---|
| `4` | 0.5 pt | `10` |

(`w:sz` is in eighths of a point; 1 pt = 20 twips, so twips = `sz × 2.5`.) AUTO-width tables
(`type="auto"`) are **not** normalized (Word leaves them bare), and a table that already declares
`w:tblCellMar` is left untouched.

| Renderer | Fixed-width table, no `tblCellMar` |
|---|---|
| Word (compare output) | inserts `tblCellMar`/`tblInd` = border width (hairline) → cells fill the fixed columns |
| LibreOffice (our un-normalized output) | applies its own ≈108-twip default → cell text shifts right, table ghosts |
| Docxodus (after fix) | inserts border-width inset like Word → renders where Word's does |

### The fix

`WordCompareTableNormalizer.NormalizeAll` runs as a single-owner post-pass over the assembled body blocks
in `IrMarkupRenderer.Render`. For each `w:tbl` whose `w:tblW` is `dxa`, that has a derivable border width
and no declared `w:tblCellMar`, it inserts `w:tblCellMar` (left/right) and `w:tblInd` at the border-width
inset, in CT_TblPrBase schema order. This mirrors the docDefaults backfill (`WordStockDocDefaults`) — the
engine's job is to reproduce Word's compare output, so its tables must land where Word's do. The inset is
a table property, not tracked-changes markup, so the `accept ≡ right` / `reject ≡ left` contract (verified
at body-text level) is unaffected.

**Known limitation:** on a pathological "repaired" document whose oracle itself normalized only *some* of
its tables, the rule over-fires on the un-normalized ones (no principled source condition distinguishes
them). The affected document is a deep residual on other grounds; the net effect across the corpus is a
strong improvement with no table rendered worse than the un-normalized default.

### Relevant code

- `Docxodus/Ir/Diff/WordCompareTableNormalizer.cs` — the normalization rule.
- `Docxodus/Ir/Diff/IrMarkupRenderer.cs` — `Render` invokes the post-pass after block assembly.

### Tests

`Docxodus.Tests/DocxDiffTableCellMarginBackfillTests.cs` — border-width backfill, border-width tracking
(`sz="8"` → 20), auto-width untouched, existing-margin untouched, no-border untouched.

---

## Tracked changes: author-color leak when input revisions are preserved

**Status:** Fixed (DocxDiff — `DocxDiffSettings.NormalizeRevisionAuthors`, opt-in)

### Symptom

A DocxDiff redline of a document whose source (the revised side) already carried tracked changes —
e.g. Google-Docs "suggestions" authored by *Online User*, or another reviewer's edits — rendered (in
LibreOffice) with that preserved content in a DIFFERENT COLOR from the fresh compare markup, while
Microsoft Word's oracle redline of the same pair was a single color. On a page dominated by such
content this cost large amounts of pixel-diff energy even though the tracked-changes markup was
structurally faithful.

### The corner case

LibreOffice colors tracked changes **by author**: each distinct `w:author` gets its own color. With
`PreserveInputRevisions` on (an opt-in on the raw `DocxDiff` API; the `DocxCompare` front door leaves it
off), the inputs' own revisions
ride through into the output under their **original** author, and the fresh compare revisions use the
engine's own author — so the output carries **two** authors and renders in **two** colors. Word's
compare output, for these documents, is single-author (its oracle carries one `w:author`), hence one
color.

Empirically across the corpus this is common but not universal — of 42 documents where our output
leaked a second author, 29 had single-author oracles (Word collapsed to one) and 13 had genuinely
multi-author oracles (Word kept more than one). Because there is no reliable structural signal that
distinguishes "Word collapsed" from "Word kept" on the input alone, this is exposed as an **opt-in
setting**, not default behavior.

| Renderer | Source with a preserved foreign-author revision |
|---|---|
| Word (compare output, these docs) | single `w:author` → one color |
| Docxodus (`PreserveInputRevisions` on) | fresh author + preserved author → two colors |
| Docxodus (`NormalizeRevisionAuthors` on) | all revision authors collapsed to one → one color |

### The fix

`IrMarkupRenderer.NormalizeRevisionAuthors` runs as a byte→byte post-pass over the rendered document
when the flag is set: it stamps `settings.AuthorForRevisions` onto the `w:author` of every
tracked-revision element (`w:ins`/`w:del`/`w:moveFrom`/`w:moveTo` and the `*Change` markers) across the
wordprocessing story parts (document, headers, footers, footnotes, endnotes, comments-part revisions),
leaving `w:comment` authors untouched (a comment is not a tracked change) and non-wordprocessing parts
(charts, customXml, SmartArt) alone. Author is presentation metadata, so revision structure — and the
`accept ≡ right` / `reject ≡ left` contract — is unaffected.

This normalizes the **render**, not the markup semantics — it matches how an author-coloring renderer
displays Word's single-author output, and it is a net-positive approximation (correct for the 29
single-oracle docs, a no-op-or-neutral change for the 13 multi-oracle docs in practice) enabled via an
explicit setting. It deliberately does not touch Consolidate/N-way output, whose per-reviewer authors
are intentional.

### Relevant code

- `Docxodus/Ir/Diff/IrMarkupRenderer.cs` — `NormalizeRevisionAuthors` / `RevisionBearingParts`.
- `Docxodus/Ir/Diff/IrDiffSettings.cs`, `Docxodus/DocxDiff.cs` — the flag + public mirror.

### Tests

`Docxodus.Tests/DocxDiffAuthorNormalizationTests.cs` — collapse of a preserved foreign author to the
single author; the leak when the flag is off.

---

## Tracked changes: superseding pending text nests `w:ins > w:del`, and an author's own text just disappears

### Symptom

Two tracked whole-paragraph replacements in a row (author A, then author B) accepted to the
concatenation of both proposals (issue #786): the first replacement's `w:ins` was left live
beside the second one's `w:del` of the original text.

### What Word writes

When B deletes text that A inserted and nobody has accepted yet, Word does not remove A's
insertion and does not wrap it: it records B's deletion *inside* A's insertion, converting the
text to its deleted spelling:

```xml
<w:ins w:author="A"><w:del w:author="B"><w:r><w:delText>Hamburg</w:delText></w:r></w:del></w:ins>
<w:ins w:author="B"><w:r><w:t>Köln</w:t></w:r></w:ins>
```

Accepting everything keeps only `Köln`; rejecting everything restores the original; rejecting
only B's deletion brings `Hamburg` back as A's pending insertion. When B deletes text that **B**
inserted, Word writes nothing at all — the text simply disappears, envelope and all — because a
pending insertion its own author withdraws has no reviewable state left. The same holds for a
whole paragraph B inserted and then deleted: it is gone, not marked.

A deleted hyperlink keeps its shell and carries the deletion inside it (`w:hyperlink > w:del >
w:r`); an inserted one carries the insertion the same way. `CT_RunTrackChange` does not admit
`w:hyperlink` as a child, so `w:ins > w:hyperlink` fails schema validation.

| Engine | B replaces A's pending text | B replaces B's pending text |
|---|---|---|
| Word | `w:ins(A) > w:del(B)`; accept keeps only the latest | text removed outright |
| Docxodus before #786 | A's `w:ins` left live beside B's `w:del` of the original | same |
| Docxodus | matches Word | matches Word |

### Docxodus code

`DocxSession.DeleteInlineElementInPlace` is the single owner of the in-place deletion policy
(`ReplaceText`, `DeleteBlock`/`DeleteRange`/`DeleteSection`, content-control fills); the
stateless `RevisionProcessor` reverses `w:del` only under `w:p`, `w:hyperlink`, `w:fldSimple`,
`w:dir` and `w:bdo` (looking through `w:sdt`/`w:sdtContent`/`w:smartTag`), which is why
deletions are written per run inside their carriers and never as `w:del > w:sdt` or under
`w:customXml`.

## Office Math: revision wrappers nested inside `m:r` are invalid, and Docxodus preserves them

**Status:** Deliberate — not repaired (decided 2026-09-02, issue #642)

### The behavior

A small set of older Word documents place a tracked-revision wrapper (`w:ins`/`w:del`) *directly
inside* an Office Math run, as `m:r`'s child. The schema does not allow it: `m:r`'s content model
is `m:rPr?`, `w:rPr?`, then the text elements, and revision marking on a math run belongs inside
`w:rPr`. `OpenXmlValidator` reports exactly one finding per occurrence:

```text
[Schema] The element has invalid child element 'w:ins'.
  List of possible elements expected: <m:rPr>.
  path: /w:document[1]/w:body[1]/w:p[2]/m:oMathPara[1]/m:oMath[1]/m:r[2]
```

### Minimal reproducer

`TestFiles/WC/WC012-Math-After.docx` carries the shape verbatim:

```xml
<m:r>
  <w:ins w:id="0" w:author="Eric White" w:date="2016-04-21T19:17:00Z">
    <w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr>
    <m:t>2</m:t>
  </w:ins>
</m:r>
```

### Consumer comparison

| Consumer | Result |
|---|---|
| `OpenXmlValidator` | 1 schema finding |
| LibreOffice 25.8 | loads and renders the document without complaint |
| `DocxDiff` / `DocxCompare` | returns the input's bytes unchanged, finding included |
| `WmlComparer` (removed in v11.0.0) | emitted a repaired package — 0 findings, different bytes |

### The decision

**Docxodus preserves the shape rather than rewriting it.** Repairing it would mean silently
changing bytes a caller did not ask us to change, in a package that no consumer we can test is
actually choking on. `WmlComparer` did repair it, but only as a side effect of rebuilding the whole
document from its own model — the same whole-package reserialization that made it drop content
elsewhere (it also took this document's validator findings from 80 to 29 on unrelated markup). That
is not a property worth reconstructing deliberately.

The consequence is stated rather than implicit: an invalid input stays invalid on the way out, and
Docxodus is not a document-repair tool. A caller that needs the shape normalized should validate its
inputs before feeding them in. If a consumer is ever found that genuinely fails on this markup, the
right home for a repair is `StrictOoxmlNormalizer`, so every entry point benefits rather than only
comparison.

This also explains why `DocxCompare.CanReturnExactNoOp` is now plain byte equality. Through v10 it
deliberately refused the exact-clone shortcut for this shape, so the document would be routed
through the repairing legacy path instead. With no engine repairing it, that guard only bought a
full comparison that returned the same bytes and the same finding.

### Tests

`Docxodus.Tests/DocxCompareTests.cs` —
`ByteIdenticalMalformedMathRevision_PassesThroughUnrepaired` pins the behavior: the shortcut is
taken, and the schema finding survives on both sides.

---

## Headers/footers: a first/even part can outlive its `w:titlePg` / `w:evenAndOddHeaders` flag

### The behavior

A `w:headerReference`/`w:footerReference` of type `first` or `even` is **inert on its own**. Word
renders the first-page stories only when the governing `w:sectPr` carries `w:titlePg`, and the
even-page stories only when the settings part carries `w:evenAndOddHeaders`.

The trap is what Word does when the user turns those options back **off** in the UI ("Different
first page" / "Different odd & even pages"): it removes only the flag. The header/footer parts and
their references stay in the package. A document can therefore carry a complete set of six stories
— default/first/even for both header and footer — with neither flag set, which is exactly the shape
of `TestFiles/HC031-Complicated-Document.docx`.

### Minimal reproducer

```xml
<!-- word/document.xml — references present, no w:titlePg -->
<w:sectPr>
  <w:headerReference w:type="first"   r:id="rId15"/>
  <w:headerReference w:type="default" r:id="rId12"/>
  <!-- no <w:titlePg/> -->
</w:sectPr>
```

`word/header3.xml` (the `first` part) can hold arbitrary content and **no renderer will show it**.

### Renderer comparison

| Renderer | Result |
|---|---|
| Word | First-page header ignored; the Default header renders on page 1. |
| LibreOffice | Identical — verified by converting to PDF; the first/even content never appears. |
| Docxodus (`WmlToHtmlConverter`, pagination) | Same: `selectHeader`/`selectFooter` only pick the first/even story when the flags resolve. |

So all three agree. The hazard is not a rendering divergence — it is that **writing content into
such a story appears to succeed and silently produces an invisible result**.

### Why it bites a mutation API

`DocxSession.SetHeaderText`/`SetFooterText` set the flags as a side effect of writing content, so
authoring a story from scratch is fine. But a caller that *edits an existing* first/even story —
via `ReplaceText`/`ApplyFormat` on the story's paragraph anchor, which is what an anchor-addressed
editor does — never goes through that path, and the flag is never added. The saved file then
contains the user's text in the right part, rendered nowhere.

### Relevant code

- `Docxodus/DocxSession.cs` — `EnsureHeaderFooterVisible` (the section-level operation that sets
  the flags independently of a content write); `SetHeaderFooterText` (the create-time path).
- `Docxodus/WordprocessingMLUtil.cs` — `EnsureEvenAndOddHeaders` (inserts the settings child at its
  CT_Settings schema slot; see the settings-ordering note above).
- `npm/src/editor-headerfooter.ts` — the editor band calls the op whenever `first`/`even` is
  selected, because selecting that kind *is* the user asking for a different first/even page.

### The second-order surprise (worth surfacing in a UI)

Turning either flag on means those pages stop inheriting the Default stories **entirely**, and
`w:evenAndOddHeaders` is document-global and governs footers as well as headers. A section with a
populated Default footer but an empty even footer therefore shows *no footer at all* on even
pages, and enabling `w:titlePg` with an empty first-page footer leaves page 1 without one. This is
spec-correct and reproduces identically in Word and LibreOffice; the editor's header/footer bands
show an inline note for both cases.

### The render side of the same rule

The flag governs **reading** too, and our paginated renderer originally got it wrong in the
opposite direction: `RenderPaginatedHeaderFooterRegistry` gated the *first* stories on
`w:titlePg` but emitted the *even* stories whenever a `w:type="even"` reference existed, with no
check on `w:evenAndOddHeaders`.

| Renderer | Even-page footer, reference present, `w:evenAndOddHeaders` absent |
|----------|------------------------------------------------------------------|
| Word | Default story |
| LibreOffice 25.8 | Default story |
| Docxodus (before) | **Even story** |
| Docxodus (after) | Default story |

Found by smoke-testing the NVCA model certificate of incorporation
(`https://nvca.org/wp-content/uploads/2025/10/NVCA-Model-COI-10-1-2025.docx`), a real filing
template that carries three `w:type="even"` footer references with the flag absent. Its leftover
even footer reads `DRAFT` and carries no `PAGE` field, so the paginated view showed `DRAFT` — and
therefore no page number at all — on every even page, where LibreOffice showed
`Last Updated October 2025` and the roman-numeral page number.

The fix mirrors the existing `hasTitlePage` gate with a `hasEvenAndOddHeaders` one, so both stories
are governed by their own flag in the same place.

### Tests

`Docxodus.Tests/DocxSessionTests.cs` — `DS268` (flags set for pre-existing stories, idempotent,
lands in the section that carries the reference rather than merely the trailing `sectPr`);
`Docxodus.Tests/PaginatedHeaderFooterGatingTests.cs` — `PHF001`/`PHF002` (even stories follow
`w:evenAndOddHeaders`) and `PHF003` (first stories follow `w:titlePg`, pinned so the two rules
cannot drift apart); `npm/tests/editor-headerfooter.spec.ts` — the end-to-end assertion over the
saved package.

---

## Endnotes: the default `w:numFmt` is lowerRoman, not decimal

### The behavior

Footnote and endnote markers do **not** share a default numbering format. With no `w:numFmt`
declared anywhere — no `w:footnotePr`/`w:endnotePr` in the settings part, none in any `w:sectPr` —
Word and LibreOffice number footnotes `1, 2, 3…` but endnotes `i, ii, iii…` (lowercase roman).
ECMA-376 pins this: `w:numFmt`'s default is `decimal` in a footnote context (§17.11.17/§17.11.18)
but `lowerRoman` in an endnote context (§17.11.17/§17.11.19). A typical Word-authored package
carries a `w:endnotePr` in settings.xml that declares only the two reserved separator notes and
**no** `w:numFmt` at all, so the spec default is what actually renders.

Precedence when the format IS declared: a `w:numFmt` inside a section's `w:sectPr/w:endnotePr`
overrides the document-wide declaration in settings.xml for that section.

### Minimal reproducer

```xml
<!-- word/settings.xml — Word's usual shape: endnotePr present, numFmt absent -->
<w:endnotePr>
  <w:endnote w:id="-1"/>
  <w:endnote w:id="0"/>
</w:endnotePr>
```

Cite one endnote from the body (`TestFiles/WC/WC036-Endnote-With-Table-Before.docx` is exactly
this shape).

### Renderer comparison

| Renderer | Endnote marker | Footnote marker |
|----------|----------------|-----------------|
| Word | `i` | `1` |
| LibreOffice | `i` | `1` |
| Docxodus (before fix) | `1` | `1` |
| Docxodus (issue #414 fix) | `i` | `1` |

### Relevant code

`WmlToHtmlConverter.GetNoteNumberFormat` resolves the effective token (sectPr-level `notePr` →
settings-part `notePr` → spec default), `FormatNoteNumber` renders the glyph via
`ListItemTextGetter_Default.GetListItemText`, and the footnotes/endnotes section `<ol>` carries
the matching CSS `list-style-type` (`NoteListStyleType`).

## Layout: Word's line breaking and table sizing have no matching CSS default

### The behavior

Two of Word's layout invariants are the OPPOSITE of what CSS does by default, so an HTML
rendering of a DOCX gets them wrong unless the converter says otherwise.

1. **A word wider than its column is broken, not overflowed.** Word and LibreOffice put as much
   of an over-long word on the line as fits and break the rest. CSS's initial
   `overflow-wrap: normal` never breaks inside a word, so it runs past the margin instead.

2. **A table never exceeds the text column.** Word's table layout — fixed or AutoFit — keeps the
   table inside the column, narrowing columns when it must. CSS's `table-layout: auto` does the
   reverse: a cell's widest unbreakable word is a min-content floor, and the table is GROWN until
   that floor is satisfied, container be damned. A table box is not shrinkable below it.

The two compound. Enlarging a run inside a fixed-width cell raises the cell's min-content, which
widens the table, which overflows the page — the visible symptom being document text painted
outside the sheet and clipped by the window.

### Minimal reproducer

```xml
<w:tbl>
  <w:tblPr>
    <w:tblW w:w="9936" w:type="dxa"/>   <!-- 496.8pt: a full-width US Letter table -->
  </w:tblPr>
  <w:tblGrid><w:gridCol w:w="9936"/></w:tblGrid>
  <w:tr><w:tc><w:p><w:r>
    <w:rPr><w:sz w:val="132"/></w:rPr>   <!-- 66pt -->
    <w:t>Edit this document.</w:t>
  </w:r></w:p></w:tc></w:tr>
</w:tbl>
```

Rendered into a 354pt-wide container (a phone):

| Renderer | Result |
|----------|--------|
| Word | Table at the text column width; "document." breaks or wraps inside it |
| LibreOffice | Same; the page is zoomed to fit the window, never reflowed narrower |
| Docxodus (before) | `<table>` grown to ~462pt by the cell's min-content; text clipped by the window |
| Docxodus (after) | Table at the column width, word wrapped, page zoomed to fit |

### Analysis

CSS has no single property for "lay out like a word processor". The behaviors have to be
assembled:

- `overflow-wrap: break-word` gives Word's line breaking, but deliberately does NOT change
  intrinsic sizing — which is why it alone does not stop the table from growing. (Chrome's UA
  stylesheet sets it on `[contenteditable]`, which is why an editable paragraph appears to break
  correctly while the table around it still overflows.)
- `overflow-wrap: anywhere` DOES lower min-content, so it is the right rule for table cells.
- `table-layout: fixed` (with the column widths in a `colgroup`) is the analogue of Word's
  `w:tblLayout` fixed: authored widths become binding and content wraps inside them.
- `max-width: 100%` enforces "a table never exceeds the text column" for the AutoFit case.

None of it helps if the container is the wrong width to begin with: the text column must come
from `w:sectPr`, and a window narrower than the page must ZOOM, the way every word processor
does, rather than reflow.

### Relevant code

`Docxodus/WmlToHtmlConverter.cs` — `GenerateDocumentLayoutCss` (always emitted, ahead of the
caller's `GeneralCss` so a consumer can still override), `IsFixedLayoutTable`/`CreateColGroup`
(`ProcessTable`), and `CreateSectionDivs`, which stamps section geometry in every render mode
rather than only under `PaginationMode.Paginated`. `npm/src/page-geometry.ts` reads that geometry;
`npm/src/viewport.ts` (`DocumentViewport`) applies the column width and the fit-to-width zoom.

### Tests

`Docxodus.Tests/HtmlConverterTests.cs` — `HC056` (layout CSS always emitted), `HC057` (section
geometry outside Paginated mode), `HC058`/`HC059` (fixed vs AutoFit table layout);
`npm/tests/editor-page-geometry.spec.ts` — the end-to-end assertion that a 66pt heading on a
390px viewport leaves nothing overflowing the sheet.

---

## Package Output

### Misleading Deflate Hints Cause Compression Loss

**Status:** Fixed (August 2026)<br>
**Issue:** #331<br>
**Test:** `Docxodus.Tests/PackageCompressionTests.cs` (PKG331–PKG334)

#### The problem

Word-authored OPC packages commonly set bits 1–2 of each ZIP entry's general-purpose flag to the
"superfast" deflate hint, even when the existing compressed bytes have a high compression ratio.
That hint is not evidence of how efficiently the current bytes were actually compressed.

.NET 10's `ZipArchive` update-mode constructor maps those bits back to a compression policy for a
future rewrite: normal → `Optimal`, maximum → `SmallestSize`, and both fast/superfast → `Fastest`.
`System.IO.Packaging` preserves unchanged entries efficiently, but an XML part opened for writing
inherits the source entry's policy. A normal Open XML save can therefore keep every unchanged
binary part intact while recompressing changed XML with `Fastest`.

On `HC031-Complicated-Document.docx`, a representative 25-part fixture, the pre-fix session save
showed the characteristic disproportion:

| Scope | Before save | Pre-fix output | Uncompressed change |
|---|---:|---:|---:|
| `word/document.xml` compressed bytes | 19,290 | 24,219 | +742 |
| Whole package bytes | 42,336 | 51,491 | about +800 total |

The package grew by 9 KB even though the actual XML grew by less than 1 KB. Recompressing the exact
same output payloads with an explicit policy produced 37,101 bytes at `Optimal` and 36,724 bytes at
`SmallestSize`, isolating compression selection as the cause.

#### Output policy

`ZipPackageOutputNormalizer` runs only after the owning `Package`/`OpenXmlPackage` has finished
writing an output. Byte-exact clone and no-op comparison paths return the cloned package directly,
so they neither change ZIP representation nor pay the finalization cost. For modified output, the
normalizer builds a fresh archive in one streaming pass and:

1. copies every entry payload without parsing or reserializing it, preserving content types,
   relationship parts, signature parts, macros, and all other OPC semantics;
2. preserves entries as stored when their completed source representation got no benefit from
   compression (the normal case for JPEG/PNG and similar media);
3. uses `CompressionLevel.Optimal` for frequently written package markup, balancing size and save
   latency, and `SmallestSize` for compressible binary assets such as embedded fonts;
4. preserves entry names, order, timestamps, comments, and existing external attributes; and
5. assigns sane Unix permissions to zero-attribute entries, subsuming the issue #302 output pass
   so the archive is rewritten only once.

This is deliberately an output-boundary policy, not a byte-header patch or reflection workaround.
Changing ZIP flags on a live package would put its archive bookkeeping out of sync, while setting
`OpenXmlPackage.CompressionOption` controls only newly created parts and cannot repair existing
parts rewritten in update mode.

#### Performance tradeoff

Finalization performs one sequential inflate/copy/deflate pass and holds the source and destination
archives in memory. That costs more CPU than returning the update-mode ZIP directly, but the
latency-sensitive XML path uses `Optimal`, stored media avoids compression work, and the pass
replaces the previous separate Unix-metadata rewrite. The cost is bounded to final output
production; editing, projection, undo, and intermediate operations are unchanged. This policy
favors storage and transfer efficiency for batch-produced documents without paying maximum
compression cost on every XML save.

### Content cloned into another part loses its namespace declarations

**Status:** Fixed<br>
**Issue:** #836<br>
**Tests:** `Docxodus.Tests/DocxDiffImportedNamespaceTests.cs`, `Docxodus.Tests/PartNamespacesTests.cs`

#### The problem

Word declares every namespace a part uses on the part's root element and lists its extension
prefixes in the root's `mc:Ignorable`. A comparison copies content from the revised document into
parts that came from the original: list definitions into `numbering.xml`, styles into
`styles.xml`, paragraphs into `document.xml`, notes and comments into their parts. Copying uses
LINQ to XML clones, which keep each element's and attribute's namespace but not the declarations
on the ancestors they were cloned from. When the original's part root declares only `w:` (common
for documents produced by a generator rather than by Word), the output part contains names and
prefix lists that nothing in scope declares:

```xml
<!-- original numbering.xml declares only w:; the list definition is copied from a Word document -->
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="1" p3:restartNumberingAfterBreak="0"
                 xmlns:p3="http://schemas.microsoft.com/office/word/2012/wordml">
    <w:lvl w:ilvl="1">
      <AlternateContent xmlns="http://schemas.openxmlformats.org/markup-compatibility/2006">
        <Choice Requires="w14"><w:numFmt w:val="custom" w:format="001, 002, 003, ..."/></Choice>
        <Fallback><w:numFmt w:val="decimal"/></Fallback>
      </AlternateContent>
```

The XML writer invents a prefix for each undeclared attribute namespace (`p3:`) and redeclares the
default namespace for each undeclared element namespace. Both are well-formed, but neither is in
`mc:Ignorable`, so a consumer that does not understand the namespace must reject the content
rather than skip it. `Requires="w14"` is plain text to LINQ to XML, and here nothing declares
`w14`. Markup Compatibility (ECMA-376 Part 3) requires a `Requires` prefix to resolve.

| Consumer | Output before the fix | Output after the fix |
|---|---|---|
| Open XML SDK 3.5 validator (Office 2019) | `Sch_UndeclaredAttribute` on the `w15`/`w16cid` attributes, `MC_InvalidRequiresAttribute`, and `Sch_InvalidElementContentExpectingComplex` for a `w14` element in a style's `w:rPr`; the inputs validate clean | no errors |
| LibreOffice 24.2 (text export) | renders the `mc:Choice` branch (`001`); it does not require the prefix to resolve | same |
| Word | not tested here | root declarations and `mc:Ignorable` match the shape of Word-authored parts |

#### The fix

`PartNamespaces` records the namespace declarations and `mc:Ignorable` prefixes on the part roots
of the documents the output is built from. At the end of `IrMarkupRenderer.Render` and
`IrCompositeMarkupRenderer.Render`, it visits every part the render loaded. On each part's root
it declares every namespace that some name uses without a declaration in scope, and every prefix
that an `mc:` attribute or `mc:Choice/@Requires` lists but nothing declares, using the prefix the
input used. It then adds to `mc:Ignorable` each namespace that the part uses and an input listed
there. It never rebinds a prefix the root already uses. A part it does not change is not
rewritten.

#### Relevant code

- `Docxodus/Internal/PartNamespaces.cs` — the source declarations and the pass.
- `Docxodus/Ir/Diff/IrMarkupRenderer.cs`, `Docxodus/Ir/Diff/IrCompositeMarkupRenderer.cs` — the
  single call at the end of each render.

### Comparison output: shared bookmark ids, dangling note references, doubled singleton parts

**Status:** Fixed<br>
**Issue:** #840<br>
**Tests:** `Docxodus.Tests/DocxDiffPackageIntegrityTests.cs`, `Docxodus.Tests/BookmarkIdsTests.cs`

A comparison output combines content from two packages that each number their own bookmarks,
notes and relationships. These are three places where those numbers met.

#### Bookmark ids

Both documents number bookmarks from 0. When content from each survives, as a deleted paragraph
next to an inserted table, both copies of id 0 appear in the output:

```xml
<w:p><w:del ...><w:bookmarkStart w:id="0" w:name="LeftText"/>...<w:bookmarkEnd w:id="0"/></w:del></w:p>
<w:tbl>...<w:tr><w:trPr><w:ins .../></w:trPr>
  <w:bookmarkStart w:id="0" w:name="RightRow"/><w:tc>...</w:tc><w:bookmarkEnd w:id="0"/></w:tr></w:tbl>
```

The Open XML SDK validator reports `Sem_UniqueAttributeValue` on both `w:bookmarkStart` and
`w:bookmarkEnd`. It does **not** check that bookmark *names* are unique, so two copies of
`_GoBack` with different ids validate. Word's behaviour with repeated ids was not tested here.

The two-way renderer already made ids unique, but only for run-level markers in the body. Row-
and table-level markers (`w:tr`/`w:tbl` children), markers inside `m:oMath`, other stories, and
the consolidate renderer were not covered. `BookmarkIds` now renumbers across every story. The id
is an end's only link to its start, so the pass pairs markers before it renumbers them. An end
closes the earliest open start of its id on the same revision side: inside `w:del`/`w:moveFrom`,
a deleted paragraph mark, row or cell; inside `w:ins`/`w:moveTo`, an inserted one; or unchanged.
If no start on its side is open, it closes an unchanged start (or any start, when the end itself
is unchanged): a bookmark only the revised document has can open in a paragraph both documents
share and close in an inserted one, while the original's bookmark of the same id runs between
two deleted paragraphs. Only then does it close the earliest open start, so a deleted end never
closes an inserted start while a better one is open. An end with no start left is dropped.

Of the bookmarks sharing an id, one the comparison left untouched keeps it. ECMA-376 makes a
bookmark id unique per document, while the validator only checks per part, so the pass works
across stories; but two untouched bookmarks in different stories already shared their id in the
source and are left alone. A bookmark inside math or a drawing is renumbered like any other.
That content is compared by hashing it whole, and the hash now leaves out bookmark ids
(`IrHasher`), so a replaced equation whose bookmark was renumbered still reads as the original
after reject.

A bookmark only one document has can also straddle unchanged text, as above: its start bare in
the shared paragraph, its end inside `w:ins`. Reject kept the start and dropped the end.
`NormalizeBookmarks` now wraps the bare marker in the revision its partner is in.

#### Note references

The renumber pass that gives notes body-order ids took deleted notes from a queue for deleted
references. An original note paired with a revised note, while its own reference is deleted, is
not a deleted note. The queue was then one short, each later deleted reference took the next
note's definition, and the last had none. Deleted references now look up their note by their
(original) id, following a matched note to the revised id it was renumbered to. Separately, a
revised note with no blocks (`<w:footnote w:id="1"/>`, as in `WC/WC064-Footnote-Mod.docx`)
produced no note diff at all, so an output from an original without a footnotes part had a
reference and no part. Such a note now yields an empty insert, and the output carries it as the
input did.

A reference inside `w:moveFrom` is a deleted reference as much as one inside `w:del`; treating it
as live gave two swapped paragraphs' notes the same id. And the body and the notes can pair
differently. The body pairs references by position, so when both footnoted paragraphs are
rewritten both references stay equal. The notes pair by content, so the original's second note
can match the revised first, leaving the original's first note deleted and named by no reference.
It kept its original id, which the renumbered live notes now used (`Sem_UniqueAttributeValue` in
`footnotes.xml`). An unreferenced note whose id is taken now gets a fresh one. The equal
reference still names the revised note in both views, so after reject the first reference reads
the original's second note; making reject name the original's first note would take the body
diff and the note diff agreeing on which references correspond, which they do not today.

#### Doubled singleton relationships

The package may relate the main document part to only one styles, settings, numbering,
fontTable, theme, webSettings, footnotes, endnotes or comments part. The validator reports a
second one as `Pkg_OnlyOnePartAllowed`; with two `endnotes` relationships,
`WordprocessingDocument.Open(...).MainDocumentPart.EndnotesPart` throws "Sequence contains more
than one element". A duplicated settings part also brings `Sem_MissingReferenceElement` for the
separator notes it names.

The renderer registers each element it clones from the revised document, and after assembly
imports the parts those elements' relationship ids name from the revised package. Two paths put
original content inside a registered element. A paired table row was registered only after the
original's deleted cells and struck runs were added to it. A deleted paragraph's runs were fused
into an inserted paragraph that had already been registered (the shared-paragraph-mark shape
Word uses when a paragraph is replaced). A deleted picture's `r:embed="rId9"` was then resolved
in the revised package. When `rId9` named the revised document's endnotes, webSettings, theme
or styles part, that part was copied in under a second relationship of its type, and the picture
pointed at it. A row now registers only the pieces it clones from the revised document, and a
paragraph that is about to receive fused original runs hands its registration to its current
children first (`RenderState.KeepMediaRegistrationToCurrentChildren`).

Separately, `NormalizeBookmarks` re-closed a run-level start whose end sat outside a paragraph
(after a table's last row) with a synthetic end, repeating the id. It now counts an end wherever
it sits.

| Consumer | Before | After |
|---|---|---|
| Open XML SDK 3.5 validator (Office 2019) | `Sem_UniqueAttributeValue` on bookmark ids | no new errors |
| Open XML SDK `EndnotesPart` accessor | throws "Sequence contains more than one element" | returns the part |
| LibreOffice / Word | not verified | — |

#### Relevant code

- `Docxodus/Internal/BookmarkIds.cs` — pairing and renumbering; called from
  `IrMarkupRenderer.NormalizeBookmarks` and at the end of `IrCompositeMarkupRenderer.Render`.
- `IrMarkupRenderer.RenumberNoteIds` / `ReIdMatchedNotes`; `IrEditScriptBuilder.BuildOneStore`;
  `IrCompositeMarkupRenderer.ApplyCompositeNoteDiffs`.
- `IrMarkupRenderer.RenderModifyRow` — registers only right-sourced clones;
  `IrMarkupRenderer.EmitGapArranged` — narrows a registration before fusing original runs in.

---

### Table, row and property children out of schema order, and rowless tables

**Status:** Fixed<br>
**Issue:** #837<br>
**Tests:** `Docxodus.Tests/DocxDiffSchemaOrderTests.cs`, `Docxodus.Tests/SchemaChildOrderTests.cs`

#### The problem

The children of `w:tbl` (CT_Tbl), `w:tr` (CT_Row), `w:pPr` (CT_PPr) and `w:rPr` (CT_RPr /
CT_ParaRPr) form fixed sequences. The comparison renderer inserts and re-parents many of these
children — revision markers, `*PrChange` elements, whole shells — and each insertion site chose
its own position. One of them was wrong: marking a whole row inserted or deleted creates a
`w:trPr` for the `w:ins`/`w:del` marker with `AddFirst`, which lands ahead of a `w:tblPrEx` the row
already carries:

```xml
<!-- input row: property exceptions, no row properties -->
<w:tr><w:tblPrEx><w:tblBorders>…</w:tblBorders></w:tblPrEx><w:tc>…</w:tc></w:tr>
<!-- output before the fix: CT_Row requires tblPrEx, then trPr -->
<w:tr><w:trPr><w:del …/></w:trPr><w:tblPrEx>…</w:tblPrEx><w:tc>…</w:tc></w:tr>
```

The Word 2010 run properties (`w14:ligatures`, `w14:textFill`, …) have a slot of their own: the
Open XML SDK validator accepts them only after every WordprocessingML run property and before
`w:rPrChange` (`<w:b/><w14:ligatures/><w:rPrChange/>` validates; `<w:b/><w:rPrChange/><w14:ligatures/>`
and `<w14:textFill/><w:lang/>` do not). The `Order_rPr` table in `PtOpenXmlUtil.cs` ranks them
between `w:sz` and `w:szCs`, which the validator rejects, so that table's extension entries are not
used for this purpose.

Separately, a table whose every row is tracked-deleted and whose text was moved elsewhere
(`w:moveFrom` inside the deleted rows — Word writes this when a table's content is cut and pasted
with Track Changes on) accepted to a table with no rows. `RevisionProcessor` removes such rows in its
move passes — `RemoveRowsLeftEmptyByMoveFrom` drops a row whose cells the move emptied, and
`AcceptMoveFromRanges` drops every element wholly inside a `w:moveFromRange` — before
`AcceptAllOtherRevisionsTransform`, so the rows were already gone when the rule that drops a
wholly-deleted table looked for them. When the range opens before the table and closes as its last
child, the range also swallows `w:tblPr` and `w:tblGrid`, leaving an entirely empty `<w:tbl/>`:

```xml
<w:p>…<w:moveFromRangeStart w:id="10" w:author="A" w:name="move1"/></w:p>
<w:tbl><w:tblPr>…</w:tblPr><w:tblGrid>…</w:tblGrid>
  <w:tr><w:trPr><w:del …/></w:trPr><w:tc>…<w:moveFrom …>…</w:moveFrom>…</w:tc></w:tr>
  <w:moveFromRangeEnd w:id="10"/>
</w:tbl>
```

`DocxCompare.Compare` compares the accepted view of its inputs, so the shells reached the redline:
`TestFiles/RA001-Tracked-Revisions-01.docx` against `-02.docx` produced one table with `w:tblPr` and
`w:tblGrid` and no rows, and one entirely empty `<w:tbl/>`.

| Consumer | Output before the fix | Output after the fix |
|---|---|---|
| Open XML SDK 3.5 validator (Office 2019) | `Sch_UnexpectedElementContentExpectingComplex` on `w:tblPrEx`; `Sch_IncompleteContentExpectingComplex` on the empty `w:tbl`. A shell with `w:tblPr`/`w:tblGrid` and no rows passes (CT_Tbl lets the row group be empty) | no errors, no rowless table |
| Word | not verified here; a table with no rows is a known trigger for the "unreadable content" prompt | — |
| LibreOffice | not verified | — |

#### The fix

`WordprocessingMLUtil.OrderChildrenPerSchema` runs once at the end of `IrMarkupRenderer.Render` and
`IrCompositeMarkupRenderer.Render`, on every part the render loaded. It puts the children of each
`w:tbl`, `w:tr`, `w:pPr` and `w:rPr` into schema sequence using the existing `Order_pPr`/`Order_rPr`
tables and two small ones for CT_Tbl and CT_Row. It only moves nodes: a child the table does not
rank (a bookmark, a comment, whitespace) travels with the ranked child before it, Word 2010 run
properties go to their slot before `w:rPrChange` in their own schema sequence (`glow`, `shadow`,
`reflection`, `textOutline`, `textFill`, `scene3d`, `props3d`, `ligatures`, `numForm`, `numSpacing`,
`stylisticSets`, `cntxtAlts`), equal ranks keep their order, and a container already in order — and
a part with none out of order — is not rewritten.
`RevisionProcessor.AcceptRevisionsForPart` marks each table that has rows before accepting and, once the
revision transforms have run, removes each marked table left with none — whichever pass removed its
rows. A table that arrived without rows is left as it was.

Removing the shell has one consequence under `PreserveInputRevisions`: the renderer pairs each
accepted block with its original by walking both bodies in step, and the empty shell used to pair
with the original table. Without it, the walk met the original table where the accepted body had the
next paragraph and stopped, and every later block lost its preserved revisions. The walk now steps
over an original table that `RevisionProcessor.AcceptRemovesTable` says accepting removes. That
check accepts a copy of the table on its own through the same pipeline. This also fixes the older
case of a wholly-deleted table with no moves, which stopped the walk the same way. The removed table
itself is still not carried into the preserved output.

#### Relevant code

- `Docxodus/PtOpenXmlUtil.cs` — `OrderChildrenPerSchema` and its order tables.
- `Docxodus/Ir/Diff/IrMarkupRenderer.cs` (`MarkWholeRow`), `Docxodus/Ir/Diff/IrCompositeMarkupRenderer.cs`.
- `Docxodus/RevisionProcessor.cs` — `MarkTablesWithRows` / `RemoveTablesThatLostEveryRow`,
  `AcceptRemovesTable`.
- `Docxodus/Ir/Diff/IrMarkupRenderer.cs` — `AlignPreservedChildren` / `IsAcceptDiscarded`.

---

### Drawings from both compared documents share one set of drawing ids

**Status:** Fixed<br>
**Issue:** #860<br>
**Tests:** `Docxodus.Tests/DocxDiffDrawingIdTests.cs`, `Docxodus.Tests/DrawingIdsTests.cs`

#### The problem

Word numbers the drawings of a document through `wp:docPr/@id`, starting from 1, and keeps the ids
distinct across the body, headers, footers, notes and comments. A comparison that replaces one
drawing with another keeps both: the original's inside `w:del`, the revised document's inside
`w:ins`. Each keeps the id its own document gave it, so the two usually collide:

```xml
<w:p>
  <w:ins w:id="1" w:author="A"><w:r><w:drawing><wp:inline>…<wp:docPr id="1" name="Diagram 1"/>…</wp:inline></w:drawing></w:r></w:ins>
  <w:del w:id="2" w:author="A"><w:r><w:drawing><wp:inline>…<wp:docPr id="1" name="Chart 1"/>…</wp:inline></w:drawing></w:r></w:del>
</w:p>
```

VML text boxes have the same problem one level down, twice over:

- Word numbers VML shapes per document too (`<v:shape id="_x0000_s1026" o:spid="_x0000_s1026">`), so
  the deleted and the inserted text box usually carry the same `v:shape/@id`. A grouped drawing
  brings the same collision on `v:group` and `v:rect`.
- Word writes the shape type a text box uses (`<v:shapetype id="_x0000_t202" o:spt="202" …>`) once
  per part, just before the first text box, and every later text box refers to it
  (`<v:shape type="#_x0000_t202">`). The two documents each bring a copy, and the two copies usually
  differ only in `w14:anchorId`. A text box later in the document carries no definition of its own:
  it refers to whichever copy precedes it, and that copy may sit inside `w:del` or `w:ins`.

| Consumer | Duplicate `wp:docPr/@id` | Duplicate VML `id` (`v:shape`, `v:group`, `v:rect`, `v:shapetype`, …) |
|---|---|---|
| Open XML SDK 3.5 validator (Office 2019) | `Sem_UniqueAttributeValue` when the duplicates are in one part and outside `mc:AlternateContent`. The SDK's constraint is part-scoped. It did not report a duplicate planted between drawings in two `mc:Choice` branches of one header. | Part-scoped uniqueness constraints on each VML element's `id`, reported as `Sem_UniqueAttributeValue` on the element. VML that Word puts in `mc:Fallback` is not checked; a `w:pict` outside `mc:AlternateContent` is. |
| Word | not verified; Word writes ids that are unique across every story part | not verified |
| LibreOffice | not verified | not verified |

One duplicate is legitimate. Word writes the `mc:Choice` (DrawingML) and `mc:Fallback` copies of one
drawing with the same `wp:docPr/@id`, because they are the same object. A consumer reads only one
branch.

#### The fix

`DrawingIds.MakeUnique` runs once at the end of `IrMarkupRenderer.Render` and
`IrCompositeMarkupRenderer.Render`. It covers every story part: body, headers, footers, footnotes,
endnotes and comments.

- **`wp:docPr/@id`.** The pass walks every drawing in document order. The first drawing to use an id
  keeps it. A later drawing that uses the same id gets a new id above the highest id in use. Two
  drawings whose nearest common ancestor is an `mc:AlternateContent` are alternative copies of one
  object, so they keep sharing an id. Nothing in WordprocessingML refers to a `wp:docPr/@id`:
  hyperlinks (`a:hlinkClick`), charts and diagrams reach their parts through relationship ids.
  `pic:cNvPr/@id` inside a graphic is a separate id space, unique only within its `a:graphicData`.
- **VML element ids.** In each part, the first VML element to use an `id` keeps it. A later
  duplicate gets `_x0000_s` plus a number above the part's highest (for Word's `_x0000_sNNNN` form),
  or the id with a numeric suffix. Alternative copies share an id, as for drawings. An
  `o:OLEObject/@ShapeID` or `w:control/@w:shapeid` names the shape written just before it, and
  follows that shape.
- **`v:shapetype/@id`.** A definition matters only if every shape can reach one in every view of the
  document: as redlined, with all changes accepted, and with all changes rejected. The pass acts on
  an id that has more than one definition, or whose single definition some shape can lose. For
  example, a text box inserted before an existing one carries the only definition, and rejecting the
  insertion takes it away.
  - When every definition of the id is equivalent (the same elements and attribute values, ignoring
    `w14:anchorId` and namespace declarations), exactly one is kept. It is the first definition if
    no revision mark, Markup Compatibility branch, text box or tracked table row encloses it.
    Otherwise it is a copy, without `w14:anchorId`, in a new untracked `w:r/w:pict` at the start of
    the paragraph holding the first reference, or of the part's first such paragraph. That position
    survives every view, and the built-in `_x0000_t202` id is kept, so no shape is repointed.
  - When the definitions differ, or no paragraph can hold the copy, the first keeps its id. Each
    shape is bound to the nearest definition before it that is present in every view the shape is
    in: after it if none before qualifies, and failing both, the nearest before it. A later
    definition is removed if an equivalent kept definition is present wherever it is. Any other
    later definition gets a new id (`_x0000_t202_1`), and its shapes are updated to match.

  A simpler rule, where each shape uses the nearest definition before it and the later copy is
  renamed, fails in the usual case. The comparison writes the deleted copy, then the inserted copy,
  then the rest of the document. An unchanged text box further down then binds to the inserted copy
  and loses it when the changes are rejected.

#### Relevant code

- `Docxodus/Internal/DrawingIds.cs` — the pass.
- `Docxodus/Ir/Diff/IrMarkupRenderer.cs`, `Docxodus/Ir/Diff/IrCompositeMarkupRenderer.cs` — the
  single call at the end of each render.
- `Docxodus/Ir/IrReader.cs` — a `w:pict` that holds only `v:shapetype` definitions draws nothing,
  so the IR treats it as non-content, like `w:lastRenderedPageBreak`. Without that, the run the pass
  adds would count as a change, and accepting the redline would not reproduce the revised
  document's content.

---

## Paragraph Layout

### `w:lineRule="auto"` is a multiple of the FONT's line box, not of font-size

#### Symptom

Every line of body text sits too close to the next, and the error compounds down the page. It is
invisible in a single line and obvious after twenty. Where the laid-out text height *is* a
rendered dimension — a `spAutoFit` DrawingML textbox, whose height is its content — the same error
surfaces as a visibly undersized box.

#### Minimal XML reproducer

```xml
<w:pPr>
  <w:spacing w:line="259" w:lineRule="auto"/>
</w:pPr>
<w:r><w:rPr><w:rFonts w:ascii="Calibri"/><w:sz w:val="22"/></w:rPr><w:t>…</w:t></w:r>
```

`w:line="259"` with `w:lineRule="auto"` is Word's own default for documents created since Word
2013, so this is the common case, not an exotic one.

#### The corner case

ECMA-376 defines `w:lineRule="auto"` as specifying line spacing in **240ths of a line**. The
ambiguity is what "a line" means, and it is not the font size: it is the font's own single-line
height — ascent + descent + line gap, from the font's metrics.

CSS has no equivalent base. Both `line-height: 107.9%` and `line-height: 1.079` resolve against
**font-size**, i.e. the em square. For Calibri (and its metric-compatible substitute Carlito) the
font's natural line box is ≈1.22 em, so a percentage translation under-measures every line by
that ratio — about 19% — regardless of how faithfully the 259/240 arithmetic itself is done.

At 11pt (14.667px):

| Model | Line height |
|---|---:|
| `line-height: 107.9%` (percentage of em square) | 15.69px |
| 1.0792 × the font's line box (1.22 em ≈ 17.91px) | **19.33px** |
| LibreOffice, measured | **19.33px** |

Note that `w:line="240"` (exactly single) is the special case where doing nothing is right: CSS
`line-height: normal` already *is* the font's natural line box. The error only appears once the
multiplier differs from 1 — which is precisely Word's modern default.

#### Renderer comparison

| Renderer | 11pt Calibri, `w:line="259" w:lineRule="auto"` |
|---|---|
| Word | multiple of the font's single-line height |
| LibreOffice | 19.33px line advance (measured from PDF text extents at 96 DPI) |
| Docxodus (before) | 15.69px — `line-height: 107.9%` |
| Docxodus (after) | 19.33px — `line-height: normal` + `calc(1lh * 1.079)` |

#### The fix

CSS's `lh` unit is the missing base. The paragraph keeps `line-height: normal`, so `1lh` on each
direct inline child resolves to the browser's native line box for that paragraph's font, and
`calc(1lh * var(--docx-auto-line-spacing))` multiplies it. Applying it to the children rather than
the paragraph avoids a self-reference in the paragraph's own `line-height`. Nothing font-specific
is hard-coded, so the result follows whatever font actually resolves.

A child that already declares an explicit `line-height` is skipped — the compacted pieces of a
`w:br` carry `line-height: 0` precisely so they cannot contribute a line box.

#### Relevant code

- `Docxodus/WmlToHtmlConverter.cs` — `CreateStyleFromSpacing` (derives the multiplier),
  `ApplyAutomaticLineSpacingToInlineContent` (applies it), `DefineParagraphStyle` (enables it).

#### Tests

- `Docxodus.Tests/HtmlConversionOpsTests.cs` — `HCO073_PointSuffixedAutoLineSpacing_NormalizesToTwipsBeforeDerivingCss`
- `npm/tests/drawing-autofit-height.spec.ts` — the auto-fit textbox height that this drives
- `npm/tests/toc-line-geometry.spec.ts` — TOC entry line boxes

#### History

PR #372 introduced the native-line-box model but enabled it only for *empty* paragraph marks,
which is what pagination parity needed at the time; populated paragraphs kept the percentage
fallback. Issues #396 (DrawingML textbox auto-fit height) and #397 (TOC line height) were both
traced to that remaining fallback.

#### Chromium rounds `normal` to a whole pixel (issue #850)

The `lh` model above inherits one more error. Chromium sizes a `line-height: normal` line from the
font's natural line height, but rounds it to a whole CSS pixel (0.75 pt) first. Explicit lengths keep
Chromium's 1/64-pixel layout precision, so the rounding comes from `normal` alone. `calc(1lh * m)`
multiplies the rounded value.

| 11 pt Carlito (natural height 2500/2048 em) | Line pitch |
|---|---:|
| Font metrics | 13.428 pt |
| Chromium `line-height: normal` | 13.500 pt (18 px) |
| Docxodus export before #850 | 13.500 pt |
| Docxodus export after #850 | 13.43 pt (within 1/64 px) |

The 0.07 pt per line accumulates down a page. For 12 pt Liberation Sans, which is metric-compatible
with Arial, the error is 0.30 pt per line, so it moves page breaks.

| 12 pt Arial / Liberation Sans, single | Line pitch |
|---|---:|
| Word, as measured in #850 | 13.68 pt |
| Font metrics (`hhea`: 1854 + 434 + 67 per 2048) | 13.80 pt |
| Docxodus export before #850 | 13.50 pt (0.18 pt short of Word) |
| Docxodus export after #850 | 13.80 pt (0.12 pt long of Word) |

The fix removes Chromium's rounding and follows the font's own metrics. For Carlito/Calibri, all three
metric sets (`hhea`, `typo`, `win`) give the same 1.2207 em, so that is also Word's height. For Arial,
the sets disagree: `win` gives 13.41 pt and `typo` 13.06 pt. None reproduces the 13.68 pt Word figure
in #850, which has no recorded fixture to check against. The recorded Word fixture of #908 settles it:
Word's own single-spaced 12 pt Arial pitch averages 13.84 pt (13.75, 13.75 and 14.02 between four lines),
which matches the `hhea` height the export now uses.

Exact and at-least heights had a converter-side quantization of the same kind: they were formatted
with one decimal, so `w:line="253"` (12.65 pt) rendered as 12.7 pt. They now keep twentieths of a
point (`{0:0.0#}pt`). The paginated export now runs `applyUnroundedNormalLineHeights`
(`npm/src/line-metrics.ts`) before pagination. It gives every element whose computed line height is
`normal` an explicit height for its own font, measured at 1000 px, where pixel rounding is negligible.
The converter's output is unchanged, and so is the editor's paginated view. Word's baseline placement
for `auto` multiples (extra leading below the text, not split around it) is the next section.
Tests: `npm/tests/export-line-pitch.spec.ts`.

#### Word puts the extra height of `auto` multiples below the text (issue #908)

A `w:lineRule="auto"` multiple above 1 makes each line taller. Word adds all the extra height **below**
the text: a line's baseline sits where single spacing would put it. CSS splits extra leading in half,
above and below the glyphs, so before #908 every line of a 1.15 or double-spaced paragraph sat half the
extra lower than Word's.

Reproducer: `npm/tests/fixtures/line-baselines.docx`, one paragraph per page at the 72 pt top margin:

```xml
<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="276" w:lineRule="auto"/></w:pPr>
  <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr>
    <w:t>CASE1 The quick brown fox…</w:t></w:r></w:p>
```

Baselines in pt from the top of the page. Word's are the text origins of its own PDF export, recorded in
`npm/tests/fixtures/line-baselines.word.json` (Word for the web, File > Export > Download as PDF):

| 11 pt Calibri | Word, line 1 | Word, lines 2–4 | Docxodus before #908, line 1 | Docxodus after #908, line 1 |
|---|---:|---|---:|---:|
| single (240) | 82.53 | 95.78, 109.28, 122.80 | 81.75 | 81.75 |
| 1.15 (276) | 82.53 | 98.03, 113.53, 128.80 | 83.25 | 81.75 |
| double (480) | 82.53 | 109.28, 136.30, 163.05 | 88.50 | 81.75 |

12 pt Arial behaves the same way: Word's first line sits at 83.28 pt at all three spacings. The export put
it at 82.50, 84.00 and 89.25 pt before the fix, and at 82.50 pt for all three after it. (The export's
values are measured in the DOM with Chromium placing baselines on whole CSS pixels, so the 1.15 shift reads
1.5 pt rather than its exact 1.0 pt.)

- **The extra height goes below the text.** The first baseline is identical at every multiple, and
  later lines advance by the multiple of the natural height (Calibri: 13.42, 15.42, 26.84 pt; the
  font's `hhea` height is 13.43 pt).
- **Whole-pixel rounding, fixed in #942.** Word sets a baseline its natural line height *L* less the
  font's descent *D* below the line's top: Calibri 1.2207 − 0.2686 = 0.952 em (10.47 pt at 11 pt), Arial
  1.1499 − 0.2119 = 0.938 em (11.26 pt at 12 pt; the ascent with the line gap above it). Chromium rounds
  the ascent and descent to whole pixels and floors half the remaining leading, so the exported first line
  sat one pixel (0.78 pt) above Word's for both fonts, every later line with it. (CSS also splits a line
  gap half above and half below the text, where Word puts it all above; Arial has a small one.) The
  first baselines at 1.15 and double now agree with single spacing to within 0.02 pt:

  | First baseline, pt from the page top | Word | Docxodus before #942 | after |
  |---|---:|---:|---:|
  | 11 pt Calibri / Carlito | 82.53 | 81.75 | 82.47 |
  | 12 pt Arial / Liberation Sans | 83.28 | 82.50 | 83.25 |

  `alignBaselinesToWord` (`npm/src/line-metrics.ts`) runs in the export once the page tree is final.
  For each paragraph whose line height is its font's natural height it computes Word's offset, *L* − *D*
  (*D* read from the canvas at 1000 px), measures Chromium's in a probe shaped like the paragraph's first
  line, and moves each direct inline child by the difference with relative positioning, so no line box
  changes. It runs after the export's clipping check on purpose: a glyph moved down stays inside its line,
  but its inline box (the font's rounded ascent plus descent, which can be taller than the line) can pass
  a page band's bottom edge by a pixel, and the check counts relatively positioned boxes. Run before
  pagination, it failed a footnote continuation as clipped. A child that sits off the baseline is not used as the probe (a raised run would otherwise have
  its raise "corrected" away), and a child holding an image, an SVG or a positioned element is not moved.
  Exact and at-least line heights are left alone, since Word's placement for them has not been recorded.
  Chromium's printed PDF snaps text origins to whole pixels, so in print the fix shows as landing on Word's
  pixel.
- **A raised run (`w:position`) grows Word's line.** CASE6 raises "raised" by 3 pt on the first line
  of a 1.15 paragraph. Word moves that line's baseline down 3 pt to make room, and every later line
  with it:

  | pt from the page top | raised run | rest of line 1 | line 2 |
  |---|---:|---:|---:|
  | Word | 82.53 | 85.53 | 101.03 |
  | Same paragraph without the raised run (CASE1) | — | 82.53 | 98.03 |
  | Docxodus before #941 | 78.72 | 81.72 | 97.16 |

  The converter raised the run with relative positioning, which does not grow the line. Since #941 it
  raises it with `vertical-align: <raise>pt`, which grows the line upward by the raise exactly as Word
  does (`ConvertRun` in `Docxodus/WmlToHtmlConverter.cs`).
- **A lowered run (negative `w:position`) grows Word's line downward.** CASE7 lowers "lowered" by 3 pt
  on the first line of a 1.15 paragraph. Word leaves that line's baseline where it was, draws the run
  3 pt below it, and makes the line 3 pt taller underneath, so every later line moves down 3 pt. That
  puts CASE7's later lines exactly where CASE6's are: both shifts grow the line by the same amount, one
  above the baseline and one below.

  | pt from the page top | rest of line 1 | lowered run | line 2 | line 3 |
  |---|---:|---:|---:|---:|
  | Word | 82.53 | 85.53 | 101.03 | 116.55 |
  | Same paragraph without the lowered run (CASE1) | 82.53 | — | 98.03 | 113.53 |

  Until #948 the converter lowered the run with relative positioning, which leaves the later lines where
  CASE1 has them. It now uses `vertical-align: -<drop>pt`, as for the raised case: the run's inline box
  moves down and the line box grows to hold it.

**Fix.** `ApplyAutomaticLineSpacingToInlineContent` (`Docxodus/WmlToHtmlConverter.cs`). Each direct inline
child keeps the multiplied `line-height: calc(1lh * m)` as before, so every line box is exactly what it was,
and is moved up by half the extra with relative positioning:
`top: min(0px, calc(1lh * (1 / m - 1) / 2))`. In `top`, `1lh` is the child's own height mN, so this is
−(m − 1)N/2, and nothing for a multiple below 1. A raised or lowered run's own `w:position` offset is added
in the same `calc()`. A child positioned some other way, or holding an image, an SVG or an absolutely
positioned descendant (a floating drawing, for which a relative child would become the containing block), is
not moved.

A first attempt instead raised each child with `vertical-align` and the single line height, which lets the
layout itself put the glyphs at the top. It changed line heights whenever a run's font differs from its
paragraph's font, because the line box is then the union of differently shaped boxes: a tab-leader
screenshot came out 3 px shorter. Relative positioning cannot change layout. Its cost is that Chromium puts
each line's baseline on a whole pixel before the offset applies, so a multiple-spaced first line can read up
to a pixel off a single-spaced one (#942).

The same measurement exposed an export bug. The paginated export's font resolver restyled only elements that
hold text, so a paragraph's own line box, which every multiple is built on, kept its requested family and
fell back to whatever the machine had installed. The resolver now also gives every rendered element whose
font matches a resolved request that face (`npm/src/font-runtime.ts`, `collectFontInventory`).

Tests: `npm/tests/export-line-baselines.spec.ts`. It lays the fixture out with subsets of Carlito and
Liberation Sans (Calibri's and Arial's metrics) served through the export's font resolver, measures
baselines in the DOM, and compares them with Word's.

### A paragraph's line box comes from the fonts on its lines, not its mark (#940)

Word sizes each line from the fonts actually on it. A paragraph mark left at the document default does not
make the paragraph's lines taller when its runs use a different, shorter font:

```xml
<!-- docDefaults: Calibri 11 pt. The mark has no w:rPr. -->
<w:p><w:pPr><w:spacing w:line="240" w:lineRule="auto"/></w:pPr>
  <w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="24"/></w:rPr>
    <w:t>CASE3 The quick brown fox…</w:t></w:r></w:p>
```

| Line pitch, 12 pt Arial runs under a default Calibri mark | |
|---|---:|
| Word (Word for the web PDF export) | 13.84 pt average (13.75, 13.75, 14.02) |
| Docxodus before #940 | 14.65 pt |
| Docxodus after #940 | 13.80 pt |

Word's baselines do not depend on the mark here. The first Word capture for #908 was of a version of
`line-baselines.docx` whose marks were plain; after the marks were given the runs' font, Word was re-captured
and gave identical baselines. `npm/tests/fixtures/line-baselines-plain-marks.docx` is the current fixture
with plain marks again, checked against that same record.

**Why Docxodus differed.** The converter took the `<p>`'s `font-family` from the mark (`PtOpenXml.FontName`
on the paragraph) and its `font-size` from the largest run. The paragraph's own line box (its CSS strut)
was then Calibri at 12 pt, 1.2207 em × 12 pt = 14.65 pt: a family-and-size pair that exists nowhere in the
document, taller than every line. Every `auto` multiple is also built on that strut (`1lh`).

**Fix.** `DefineStrutFont` (`Docxodus/WmlToHtmlConverter.cs`) takes the family and the size from the same
run, the first of the largest size, so the strut is never taller than that run's own line. A paragraph with
no runs keeps its mark's font. Whether a mark that is *taller* than its runs makes Word's last line taller
has not been recorded.

Tests: `HCO100_*` in `Docxodus.Tests/HtmlConversionOpsTests.cs`, and the plain-mark case in
`npm/tests/export-line-baselines.spec.ts`, which serves Calibri and Arial as two different faces.

### An accumulated line-spacing error can resemble a top-margin deviation

#### Symptom

In `DB012-Lists-With-Different-Numberings.docx`, later list lines once appeared about 28px lower
in Docxodus than in LibreOffice. Because the section declares a 1701-twip top margin and the first
content is a list, the shift was initially attributed to different top-margin import behavior.

#### Relevant XML

```xml
<w:sectPr>
  <w:pgSz w:w="11906" w:h="16838"/>
  <w:pgMar w:top="1701" w:right="1134" w:bottom="1701" w:left="1134"
           w:header="708" w:footer="708" w:gutter="0"/>
</w:sectPr>
```

At 96 DPI, 1701 twips is 113.4px. Glyph ink begins a few pixels below that content edge because
the ink bounds measure the font's painted pixels, not the top of its line box.

#### Renderer comparison

The fixture was rendered through Microsoft Graph's Word DOCX-to-PDF conversion and rasterized
under the benchmark's 96-DPI contract. The current Docxodus and LibreOffice artifacts use the
same page size and font-substitution contract.

| Renderer | Page (px) | First ink row | Last ink row |
|---|---:|---:|---:|
| Microsoft Graph Word conversion | 794 × 1123 | **117** | 365 |
| LibreOffice 25.8.7.3 | 794 × 1123 | **117** | 365 |
| Docxodus | 794 × 1123 | **118** | 365 |

The one-row first-glyph difference is rasterization, not layout. Docxodus and LibreOffice have
exact tolerant ink geometry (F1 1.00000), while Word independently confirms the same page-top
position. There is no top-margin deviation to emulate or renderer fix to make.

#### Analysis

The apparent 28px displacement grew with each line. That is the signature of the automatic line
spacing defect described above, not a constant page-origin offset. Once `w:lineRule="auto"` was
measured against the font's native line box, the accumulated displacement disappeared without any
change to `w:pgMar` handling. The `numbered-lists` corpus case is therefore an `environment`
residual (substituted-font rasterization), not a `reference-deviation`.

#### Evidence and tests

- `npm/tests/visual-parity/word-reference.json` records the Word page geometry, ink bounds,
  fixture hash, capture environment, and first-line measurement.
- `npm/tests/visual-parity-word-reference.spec.ts` validates that the corpus disposition cannot
  cite Word evidence unless the corresponding measurement is committed.
- `npm/tests/visual-parity/ratchet.json` records the current Docxodus/LibreOffice F1 of 1.00000.

### Cached TOC field results suppress hyperlink presentation

#### Symptom

A cached table of contents rendered blue and underlined in Docxodus while both Word and
LibreOffice rendered its entries black. An ordinary hyperlink using the same paragraph and
character styles remained blue and underlined in Word.

#### Minimal XML reproducer

```xml
<w:r><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:instrText> TOC \o "1-3" \h </w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="separate"/></w:r>
<w:hyperlink w:anchor="_Toc425251205">
  <w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr>
    <w:t>The first heading</w:t></w:r>
</w:hyperlink>
<!-- cached entries may span paragraphs -->
<w:r><w:fldChar w:fldCharType="end"/></w:r>
```

with, in `styles.xml`:

```xml
<w:style w:type="character" w:styleId="Hyperlink">
  <w:name w:val="Hyperlink"/><w:basedOn w:val="DefaultParagraphFont"/>
  <w:rPr><w:color w:val="0563C1" w:themeColor="hyperlink"/><w:u w:val="single"/></w:rPr>
</w:style>
```

#### The corner case

`w:hyperlink` alone does not imply an appearance. Here the run explicitly references the
`Hyperlink` character style, which declares blue and underline, but Word applies a higher-level
presentation rule to hyperlinks in the cached result of a complex `TOC` field. That field context,
not the `TOC1` paragraph style or the anchor name, suppresses those two properties.

The Word-reference capture of `HC022-Table-Of-Contents.docx` was rasterized at 96 DPI. In the TOC
entry region `(90,135)–(730,220)`, Word contains **0 blue pixels**, while blue content elsewhere on
the same page rules out a global color or export artifact.

| Context | Word presentation |
|---|---|
| Hyperlink inside the cached `TOC` result | Underlying TOC color; no underline |
| Ordinary hyperlink with the same `TOC1` + `Hyperlink` styles | `#0563C1`; underlined |

#### Analysis

The renderer already annotates every OOXML element with its enclosing complex-field stack before
HTML transformation. `FieldRetriever` now recognizes `TOC` and exposes whether a run belongs to its
cached result. Run styling then removes `color` and only the `underline` decoration for a
`w:hyperlink` in that context. Removing the properties instead of forcing black preserves an
intentional color supplied by the underlying TOC paragraph/run formatting.

#### Relevant code

- `Docxodus/FieldRetriever.cs` — parses `TOC` and identifies its cached result across paragraphs.
- `Docxodus/WmlToHtmlConverter.cs` — applies cached-field presentation after normal run-style
  resolution.

#### Tests

- `Docxodus.Tests/FieldRetrieverTests.cs` — pins cross-paragraph result scope and proves it ends at
  `w:fldCharType="end"`.
- `npm/tests/toc-line-geometry.spec.ts` — uses an actual cached `TOC` field plus an ordinary
  same-style hyperlink control; it pins both presentation contexts and unchanged line geometry.

#### A trap when reducing this to a generated document

A programmatically built package needs a **`DocumentSettingsPart`** for character-style resolution
to run at all. Without `word/settings.xml`, the converter still emits the style's CSS *class* on the
run but generates an **empty rule** for it, so `w:rStyle` silently loses every declared property and
the reduced case appears to pass without exercising suppression. This is the same requirement
CLAUDE.md notes for programmatic .NET test documents.

### Text before a tab that is longer than its line wraps; the tab advances from the last line

**Issue:** #891.

A paragraph whose text before a tab is longer than the line, here a long sentence followed by a
tab in a 3.25-inch table cell:

```xml
<w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr>
  <w:p>
    <w:r><w:t>the parties agree that the supplier shall deliver the services described in the schedule within thirty days</w:t></w:r>
    <w:r><w:tab/><w:t>12.50</w:t></w:r>
  </w:p>
</w:tc>
```

| Renderer | Text before the tab | Tab advance |
|---|---|---|
| Word (as reported in #891) | wraps inside the cell | to the next stop after the pen on the last line |
| LibreOffice | not measured | not measured |
| Docxodus before the fix | one unwrapped line in a `white-space: nowrap` box **9 inches** wide | to the stop after that unwrapped line |
| Docxodus after the fix | wraps in normal flow | to the next stop after the estimated last-line pen |

**Why they differed.** `CalculateSpanWidthTransform` resolves each tab from a pen position it
accumulates by measuring the runs before it. It never knew how wide the line was, so it measured
the sentence as one unwrapped line. `TransformElementsPrecedingTab` then pinned the text in an
`inline-flex` box as wide as the chosen stop. The box forced the cell, and the table, past the
page, and the paginated export failed.

**Docxodus code.** `WmlToHtmlConverter.CalculateSpanWidthTransform` now carries the line width
down the tree: the section's text width for body paragraphs (`SectionTextWidthTwips`, per section),
the cell's grid width less its margins inside a `w:tc` (`CellTextWidthTwips`), the final section's
width in headers, footers and notes, and none inside text boxes or comments. When the pen at a tab
is past the line end (less the right indent), `PenAfterWrapping` lays the words out greedily to
estimate the pen on the last line. The tab resolves from there and is marked
`pt:TabAfterWrappedText`, which makes `TransformElementsPrecedingTab` leave the text in normal flow
instead of pinning it. Text that fits keeps the pinned layout. Converter HTML for all 694
`TestFiles` fixtures is byte-identical before and after the change.

**Tests.** `Docxodus.Tests/TabAfterWrappingTextTests.cs`, and the `#891` case in
`npm/tests/export-robustness.spec.ts`.

## Paragraph layout: a paragraph holding only a floating drawing keeps its line

**Status:** Fixed (2026-10) — issue #880.

### The corner case

A paragraph whose only content is an anchored (floating, `wp:anchor`) text box, shape or picture still
occupies one line: its paragraph mark's line, at the paragraph's font size and spacing. Only the drawing
leaves the text flow. The paginated view had lifted the drawing into the page box and left the paragraph
with a zero-height line, so everything after it moved up by one line.

### Minimal XML reproducer

```xml
<w:p><w:r><w:t>Before</w:t></w:r></w:p>
<w:p><w:r><w:drawing><wp:anchor …><wp:positionV relativeFrom="page">…</wp:positionV>
  <wp:wrapNone/>… a wps text box …</wp:anchor></w:drawing></w:r></w:p>
<w:p><w:r><w:t>After</w:t></w:r></w:p>
```

### Behavior table (top of "After", points from the top of the page; 72 pt top margin)

| Middle paragraph | LibreOffice 26.2 (measured, PDF text box) | Docxodus before | Docxodus after |
|---|---|---|---|
| floating text box only | 98.93 | one line higher than the empty case | same as the empty case |
| empty `<w:p/>` | 98.93 | — | — |
| none | 85.48 | — | — |

LibreOffice puts "After" exactly where it puts it after an empty paragraph. Word lays the paragraph mark
out the same way, per the issue's report; Word was not measured here, so its column is the reported
behavior rather than a measurement.

### Analysis

The converter gives a visually empty paragraph a placeholder run (a no-break space in the paragraph
mark's run properties) so it keeps a line box. Its content test counted any `w:drawing`, including a
floating one, as content, so an anchor-only paragraph got no placeholder. Floating drawings are now
excluded from that test together with their contents (a text box's own paragraphs). An inline drawing
still counts: it is in the flow.

PageMap measurement keeps its issue #849 contract: a paragraph left with nothing but that line measures
as the union of its line and the drawings promoted out of it, so its fragment still encloses the box.

### Relevant code

- `Docxodus/WmlToHtmlConverter.cs` — `InsertAppropriateNonbreakingSpacesTransform`, `InFlowDescendantsAndSelf`.
- `npm/src/pagination.ts` — `positionDrawingAnchors` (marks the host), `measureRenderedSource`.

### Tests

`Docxodus.Tests/HtmlAnchorOnlyParagraphTests.cs` (placeholder for anchored, none for inline or for a
paragraph with text) and `npm/tests/anchor-only-paragraph-line.spec.ts` (paginated layout matches the
empty-paragraph case).

## Compare: Word marks moves below the paragraph level

### The corner case

Word's Compare marks moved text inside paragraphs, not only whole moved paragraphs, and it marks very short
spans. Docxodus reports a sub-paragraph move only when the moved text carries at least
`MoveMinimumWordCount` words (3 by default; issue #888), so Word reports moves that Docxodus draws as a
deletion and an insertion.

### What Word writes (shape observed in Word Compare output; content illustrative)

```xml
<!-- a single word moved from one table cell to another, rows left in place -->
<w:p><w:moveFromRangeStart w:id="1" w:name="move1" .../>
  <w:moveFrom w:id="2" ...><w:r><w:t>Subtotal</w:t></w:r></w:moveFrom>
  <w:moveFromRangeEnd w:id="1"/>
  <w:ins ...><w:r><w:t>Region</w:t></w:r></w:ins></w:p>

<!-- a list number moved together with the preceding paragraph mark: the range opens in the previous
     paragraph, whose mark is w:moveTo, and closes after "1. " in the next -->
<w:p><w:pPr><w:rPr><w:moveTo .../></w:rPr></w:pPr> … <w:moveToRangeStart w:id="5" w:name="move2" .../></w:p>
<w:p><w:pPr><w:rPr><w:ins .../></w:rPr></w:pPr>
  <w:moveTo ...><w:r><w:t xml:space="preserve">1. </w:t></w:r></w:moveTo><w:moveToRangeEnd w:id="5"/>
  <w:ins ...><w:r><w:t>Scope</w:t></w:r></w:ins></w:p>
```

### Behavior table

| Shape | Word | Docxodus |
|---|---|---|
| A sentence of 3+ words moved within or between paragraphs | move | move (issue #888) |
| A single word moved between table cells | move (`w:moveFrom`/`w:moveTo` in the cells) | deletion + insertion |
| A list number (`1. `) carried with the previous paragraph mark | move, the range spanning the mark | deletion + insertion |
| Identical text the two sides align differently (one side keeps it in place) | insertion | insertion |

### Analysis

Word's move detection runs over its own whole-document token stream, paragraph marks included, so a short
run of tokens that left one place and reappears in another is a move however short it is, and a move
range can open in one paragraph and close in the next. Docxodus pairs moves after a block alignment: whole
blocks by the aligner, spans inside modified paragraphs by `IrRelocationPairer`, which requires
`MoveMinimumWordCount` words so that common short phrases ("of the", a list number) are not paired
across a document. Lowering `MoveMinimumWordCount` makes shorter spans pair. A move range never spans a
paragraph mark in Docxodus output.

### Relevant code

- `Docxodus/Ir/Diff/IrRelocationPairer.cs` — span pairing and boundary slides.
- `Docxodus/Ir/Diff/IrMarkupRenderer.cs` — `BuildTokenOpContent` draws a relocated span as a move.

### Tests

`Docxodus.Tests/DocxDiffTextMoveTests.cs` (`TextBelowTheMinimumWordCount_IsNotAMove`,
`TextBelowTheMinimum_IsAMove_WhenTheMinimumIsLowered`).

## Tables: a negative `w:tblInd` pulls an over-wide table into the left margin

Word indents a table wider than the text column by a negative amount so it spreads across both
margins (issue #827):

```xml
<w:tblPr>
  <w:tblW w:w="15000" w:type="dxa"/>
  <w:tblInd w:w="-725" w:type="dxa"/>
</w:tblPr>
```

On US Letter landscape with 1" margins the text column is 12960 twips, so this 750pt table starts
36.25pt left of the column and ends inside the right margin.

| Renderer | Result |
|---|---|
| Word | Table starts 36.25pt left of the text column; fully visible, clipped only at the paper edge |
| Docxodus (before) | `margin-left: 0`; paginated content box clipped the right-hand overhang |
| Docxodus (after) | `margin-left: -36.25pt`; content box clips vertically only |

The clamp to `0` came from the upstream OpenXmlPowerTools converter, not a Docxodus decision.
The paginated content area now uses `overflow-x: visible; overflow-y: clip` — `hidden` on one
axis would force the other to `auto` and clip it anyway. The page box still clips at the paper.

`w:tblInd` is relative to the **leading** edge of the table. For a right-to-left table
(`w:bidiVisual`), the same value must become `margin-right: -36.25pt`; always setting
`margin-left` leaves its right edge at the text column. See the
[OOXML table indentation definition](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.tableindentation).
The browser regression covers both directions and zoom levels, row-split tables with horizontal
merges, actual clipping at the paper edges and body bottom, and PageMap geometry.

- Code: `Docxodus/WmlToHtmlConverter.cs` (`tblInd` in the table style; `.page-content` CSS),
  `npm/src/pagination.ts` (content area).
- Tests: `HCO099_TableIndent_UsesLeadingMargin`,
  `npm/tests/pagination-negative-table-indent.spec.ts`.

## `w:ind` has two spellings for each edge: `w:start`/`w:left` and `w:end`/`w:right`

ECMA-376 Part 1, §17.3.1.12 names a paragraph's indent edges `w:start` and `w:end`; the transitional
schema keeps the older `w:left` and `w:right` for the same edges (both are "leading" and "trailing",
so a right-to-left paragraph swaps them the same way). Word writes `w:left`/`w:right`. LibreOffice
writes `w:start`/`w:end`, so any document that has been through a LibreOffice save uses that form.

```xml
<w:p><w:pPr><w:ind w:start="1440"/></w:pPr><w:r><w:t>Indented one inch.</w:t></w:r></w:p>

<!-- numbering.xml, a list level -->
<w:pPr><w:ind w:start="720" w:hanging="360"/></w:pPr>
```

Indent measured from the left margin (LibreOffice: text x position in its exported PDF; Docxodus:
the converted paragraph's CSS). The list level is `w:left`/`w:start="720" w:hanging="360"`.

| Renderer | paragraph, `w:left="1440"` | paragraph, `w:start="1440"` | list item, `w:left` level | list item, `w:start` level |
|---|---|---|---|---|
| LibreOffice 25.8.7.3 | 1.00 in | 1.00 in | marker 0.25 in, text 0.50 in | marker 0.25 in, text 0.50 in |
| Docxodus before #894 | `margin-left: 1.00in` | `margin-left: 0` | `margin-left: 0.50in; text-indent: -0.25in` | `margin-left: 0; text-indent: -0.25in` |
| Docxodus after #894 | `margin-left: 1.00in` | `margin-left: 1.00in` | `margin-left: 0.50in; text-indent: -0.25in` | `margin-left: 0.50in; text-indent: -0.25in` |

Word was not measured; it writes `w:left`/`w:right` itself.

Before #894 a list level written with `w:start` lost its indent but kept its hanging indent, so
`text-indent: -0.25in` pulled the marker left of the paragraph's box.

**Both spellings on one element.** Nothing seen in the wild writes both, but style inheritance
used to manufacture it: `FormattingAssembler.IndMerge` merged attribute by attribute, so a style's
`w:start` survived beside the paragraph's overriding `w:left`. The merge now treats each pair
(`start`/`left`, `end`/`right`, `startChars`/`leftChars`, `endChars`/`rightChars`) as one slot, the
way it already treats `firstLine` against `hanging`. On a source element that does carry both,
every reader prefers `w:start`/`w:end`, the strict-schema name.

#### Relevant code

- `WordprocessingMLUtil.IndLeadingAttribute` / `IndTrailingAttribute` (`Docxodus/PtOpenXmlUtil.cs`):
  the one rule. The converter (`CreateStyleFromInd`, the tab layout, bordered paragraph groups),
  `FormattingAssembler.AddTabAtLeftIndent`, `DocxSession.GetFormatting` and `IndentDelta`,
  `GetListMembership` and `IrReader.MapParaFormat` all read through it.
- `FormattingAssembler.IndMerge`: the one-slot merge.

#### Tests

- `Docxodus.Tests/WmlIndStartEndTests.cs`.

---

## Theme Colors

### `w:color`/`w:fill` are a CACHE; `w:themeColor`/`w:themeFill` are the authority

#### Symptom

None, in any file Word wrote — which is exactly what makes it worth documenting. Word rewrites the
cached literal whenever it applies a theme, so the two always agree and a renderer that reads the
wrong one is indistinguishable from a correct one. The divergence only surfaces in a document whose
theme was replaced without the caches being rewritten (a template swap, a programmatic edit, or a
producer other than Word).

#### Minimal XML reproducer

A table style declaring the same accent colour twice — once as a fill, once as a border — with a
deliberately stale cache on both:

```xml
<w:tblBorders>
  <w:top w:val="single" w:sz="4" w:color="0000FF" w:themeColor="accent5" w:themeTint="99"/>
  <!-- … -->
</w:tblBorders>
<w:tblStylePr w:type="firstRow">
  <w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="FF0000" w:themeFill="accent5"/></w:tcPr>
</w:tblStylePr>
```

with `accent5` = `4472C4` in `theme1.xml`. A conforming consumer paints the header `#4472C4` and the
border `#8EAADB` (accent5 at tint `0x99`), ignoring both stale literals.

#### The corner case

ECMA-376 treats the theme reference as authoritative and the `w:color`/`w:fill` attribute as the
last computed value, retained so a consumer that cannot resolve themes still has something to draw.
Word's tint formula is `value × tint + 255 × (1 − tint)`, floored — `4472C4` at tint `0x99` (153/255)
gives `8EAADB`, and at tint `0x33` (51/255) gives `D9E2F3`, which is exactly what a real file's cache
contains.

Docxodus resolved this correctly for run colour and for shading, but **border colour read the
literal**, so one table style could derive the same accent colour from two different sources.
Fixed; the two paths now agree.

#### Renderer comparison

Measured on `HC029-Table-Merged-Cells.docx`, whose colours come entirely from the
`Grid Table 4 Accent 5` style:

| Renderer | Header fill | Band fill | Border |
|---|---|---|---|
| Word | accent5 | accent5 tint 33 | accent5 tint 99 |
| LibreOffice | `#4472C4` | `#D9E2F3` | `#8EAADB` |
| Docxodus | `#4472C4` | `#D9E2F3` | `#8EAADB` |

All three agree, because that file's cache is in sync — which is why the tracked benchmark case
could not decide the question and a generated one had to.

#### A second requirement the reduced case exposed

Word stamps every `w:tr`/`w:tc` with a `w:cnfStyle` listing the conditional formats that apply to
it. Docxodus applies table-style conditional formatting (`w:tblStylePr` for `firstRow`, `band1Horz`,
…) from those hints rather than deriving band membership from `w:tblLook` and the row index, so a
hand-authored table without them renders with **no** header or band shading at all. Real files
always carry them; a generated regression must emit them too.

#### Relevant code

- `Docxodus/WmlToHtmlConverter.cs` — `ResolveThemeColor` / `ApplyTintShade`, used by
  `CreateStyleFromShd`, run colour, and (now) `GenerateBorderStyle`.

#### Tests

- `npm/tests/table-style-color.spec.ts` — a generated table whose cached literals disagree with the
  theme, asserting the theme wins for header fill, band fill, and border colour independently.

---

## Contributing

When adding new corner cases to this document:

1. **Provide a minimal reproducer**: Include the relevant XML snippets and a description of how to reproduce
2. **Document all renderers**: Test in Word, LibreOffice, and Docxodus
3. **Reference the spec**: Link to relevant ECMA-376 sections
4. **Identify the code**: Point to the specific Docxodus files/functions involved
5. **Propose a fix**: If possible, outline how the issue might be resolved
