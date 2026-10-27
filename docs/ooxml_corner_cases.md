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
   - [Precision loss before auto-spacing (#1124)](#precision-loss-before-auto-spacing-1124)
   - [Cached TOC field results suppress hyperlink presentation](#cached-toc-field-results-suppress-hyperlink-presentation)
   - [`w:ind` has two spellings for each edge: `w:start`/`w:left` and `w:end`/`w:right`](#wind-has-two-spellings-for-each-edge-wstartwleft-and-wendwright)
5. [Theme Colors](#theme-colors)
   - [`w:color`/`w:fill` are a CACHE; `w:themeColor`/`w:themeFill` are the authority](#wcolorwfill-are-a-cache-wthemecolorwthemefill-are-the-authority)
6. [Tracked Drawings](#tracked-drawings)
   - [LibreOffice fails on tracked anchored header groups with text boxes](#libreoffice-fails-on-tracked-anchored-header-groups-with-text-boxes)
7. [Contributing](#contributing)

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

`IrMarkupRenderer.NormalizeComments` (phase B) gives each dedup clone a **fresh** `w14:paraId` (from the package-wide `CommentOps.ParaIdAllocator`, issue #1170) *and* clones the matching `commentsExtended`/`commentsIds` entry under the fresh paraId (`CloneThreadingEntryForParaId`), preserving `paraIdParent`. So the reject-side clone keeps its own threading link, exactly as the accept-side original keeps the unchanged one. Verified independently: `lo_comment_check.py` enumerates LibreOffice `Annotation` fields and asserts every reply's `ParentName` names a loaded comment — the dense fixture's Compare output reports 2 threaded replies (original + clone), both resolving.

#### `OpenXmlValidator` does NOT flag a duplicate `w14:paraId` (a second comment-threading blind spot)

Like the note-in-note blind spot above, the SDK `OpenXmlValidator` (Office2019) does **not** validate `w14:paraId` uniqueness across comment definitions — a document with two `<w:comment>/<w:p>` sharing one paraId is "schema-valid" to the validator but ambiguous to Word's/LibreOffice's threading resolver. So "no new validator errors" is **not** sufficient to prove comment-threading integrity; assert paraId/threading structurally (`DocxDiffCommentStructureTests.AnchorProjection` resolves each reply's parent through the paraId graph and checks `accept ≡ right` / `reject ≡ left` on the resolved-parent text).

#### `w14:paraId` is package-unique and lies in `[1, 0x80000000)`; a durable id has no validator bound (issue #1170)

A `w14:paraId` names one paragraph in the whole package — body, headers, footers, notes, comments and the glossary document — and the comment threading parts (`commentsExtended`, `commentsIds`) key their entries on the comment paragraphs' values. A comparison output mixes paragraphs and comments from two packages, so three things can collide: a dedup clone numbered from the comments part alone can take an id a body, header, note or glossary paragraph already holds; a paragraph or block wrapper replaced whole keeps the source's paraIds in both its deleted and its inserted copy; and a right-added comment can bring a paraId that the left uses for a different paragraph or comment.

The SDK `OpenXmlValidator` (Microsoft365) checks the value range though not uniqueness: on both `w14:paraId` and `w14:textId` it reports `MinExclusive` for `00000000` and `MaxExclusive` ("must be less than 80000000") for `80000000` and above. So an allocator that hands out "max + 1" breaks the schema once a package holds `7FFFFFFF`. It reports nothing for any `w16cid:durableId` value (`00000000`, `80000000` and `FFFFFFFF` all pass).

Docxodus: `CommentOps.ParaIdAllocator` scans every paraId-bearing part plus the threading keys and mints one past the largest id below `0x80000000`, falling back to the lowest free value; `NextDurableId` does the same below `0x7FFFFFFF`. `MergeRightCommentDefinitions` re-mints a right-added definition's paraId when the output's comments or threading already use it, re-keying its threading entries with it, and gives its `commentsIds` entry a fresh durable id when the output already uses its own. (Docxodus merges no `commentsExtensible` (`w16cex`) entries, so there is nothing to re-key there.) A last render pass, `ParagraphIds.MakeUnique`, then renames any remaining duplicate story paragraph (comment definitions and threading keys are read first and keep their values; nothing refers to a story paragraph's paraId). `w14:textId` is left alone: it is a version identifier for the paragraph's text, not an identity, so the validator's range check is the only constraint enforced on it here.

#### A reply added in the revision names its parent by the revision's paraId (issue #1312)

The redline's comment definitions are the original's. A reply the revision adds is merged in with its own `commentsExtended` entry, whose `w15:paraIdParent` is the paraId the **revised** package gives the parent comment. Nothing makes that value the original's: a tool that writes no comment paraIds leaves the original's root without one, and an editor that renumbers paraIds on save gives the same comment a different value in each version. Minimal reproducer (original, no `commentsExtended` part):

```xml
<w:comment w:id="0" w:author="Ann"><w:p><w:r><w:t>Root comment.</w:t></w:r></w:p></w:comment>
```

The revision adds `<w:comment w:id="1">` with `w14:paraId="00000002"`, gives the root `w14:paraId="00000001"`, and threads `<w15:commentEx w15:paraId="00000002" w15:paraIdParent="00000001"/>`.

| | Redline root paraId | Reply's `paraIdParent` | Reply in `ListComments` |
|---|---|---|---|
| Docxodus before | none | `00000001` (names nothing) | top-level |
| Docxodus after | freshly minted | that minted paraId | under the root |

Word was not measured. Adopting the revised paraId for the root is not a fix: when the original already gives that value to another comment, the reply would thread under that comment, or the original comment would be renamed and lose its own entry and resolved state to the root. `MergeRightCommentDefinitions` (`LinkRepliesToOriginalParents`) instead follows the parent through its comment id, which `CommentOps.AlignIds` has aligned between the inputs by author, date and text, and re-points the reply at the output definition's own paraId, minting one when it has none. A reworded root is a different comment by that identity, so the redline deletes the original's and inserts the revision's, and the reply threads under the inserted one as before.

Threading and resolution are not tracked changes, and the redline's definitions are the original's, so only the link of a reply the revision **adds** is translated. A comment the original already holds keeps the original's state: no entry in the original means open and top-level, and the redline states exactly that, even when the revision resolved it or threaded it under another original comment. This also makes a consolidation independent of reviewer order. When two original comments share author, date and text, the id alignment may pair them crosswise, so a reply to one of them lists top-level rather than under a guess. Finally, every comment with a paraId gets its own entry in each threading part the redline has (`CompleteCommentThreadingEntries`), since either input may lack a part the other brings.

#### v1 limitation: cross-document comment id / paraId collision (independent documents only)

The comment merge (`MergeRightCommentDefinitions`) and collapse assume a comment present in both sides carries the SAME `w:id` — true when the two inputs are two versions of ONE document (Word never reassigns a comment's id, so an edited doc's comment ids are stable). Two **independent** documents that each happened to assign `w:id="0"` to a DIFFERENT comment anchored on the same text, are out of v1 scope: the cross-document case can leave `accept` showing the left comment's text (the right definition is not re-id'd and merged). (A right-added comment whose `w14:paraId` collides with a left one's is handled: it takes a fresh paraId, issue #1170.) The output stays schema-valid and every reference resolves to exactly ONE comment (the `(C)` backstop guarantees that) — it is a content/threading-attribution gap, not a structural corruption, and it does not arise from the diff/review workflow the engine targets (before/after of one document). Re-id'ing right-sourced markers for a genuinely independent-document merge is a follow-on.

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

## Bookmarks and comment ranges between blocks (siblings of `w:p`)

### Symptom

A `w:bookmarkStart`/`w:bookmarkEnd` (or `w:commentRangeStart`/`End`) written directly in the body, a table cell
or a block wrapper's content vanished from a `DocxCompare`/`DocxDiff` redline, and `RevisionProcessor` accept
and reject removed it even from a revision-free document (issue #1149). Word writes a bookmark that spans whole
paragraphs, or a table, this way; the WmlComparer-era fixture `TestFiles/WC/WC-BodyBookmarks-After.docx` has 916
such markers.

### Minimal reproducer

```xml
<w:body>
  <w:bookmarkStart w:id="7" w:name="Article"/>
  <w:p><w:r><w:t>Alpha</w:t></w:r></w:p>
  <w:p><w:r><w:t>Bravo monthly</w:t></w:r></w:p>   <!-- revised: "Bravo weekly" -->
  <w:bookmarkEnd w:id="7"/>
</w:body>
```

### The corner case

The schema allows the range markers (`EG_RangeMarkupElements`) and the run-level revision wrappers `w:ins`/
`w:del` (`EG_RunLevelElements`) directly in every block stream: `CT_Body`, `CT_Tc`, `CT_SdtContentBlock`,
`CT_CustomXmlBlock`, notes, headers and footers. So a marker one document has between blocks can be tracked
right where it sits, `<w:ins ...><w:bookmarkStart .../></w:ins>` as a sibling of `w:p`, and the Open XML SDK
validator accepts that. Accept unwraps such a `w:ins` and drops such a `w:del`; reject does the reverse.

| Behaviour | Word | LibreOffice 25.8 | Docxodus |
|-----------|------|------------------|----------|
| Loads a body-level bookmark | yes (Word writes this shape) | yes (counted as a bookmark) | yes |
| Loads a body-level or cell-level `w:ins`/`w:del` around a bookmark | yes: opens with no repair prompt and lists the bookmark | yes, as a plain bookmark | writes it for a one-sided marker |
| Accept all / reject all removes a tracked bookmark | no: after Reject All Changes both the deleted-side and the inserted-side bookmark are still listed | no: LibreOffice keeps the bookmark either way, and does the same for a `w:ins`-wrapped bookmark inside a paragraph | yes (`RevisionProcessor`) |
| Accept/reject of a revision-free document keeps a body-level bookmark | not verified | yes | yes since #1149 (before: dropped) |

LibreOffice was checked with `tools/diffharness/lo/lo_bookmark_check.py` (every bookmark loads, every `REF`
resolves) and with `.uno:AcceptAllTrackedChanges`/`.uno:RejectAllTrackedChanges` on the redlines. Its bookmarks
are not revision-aware, so a one-sided bookmark survives both of its resolutions; that is not specific to
block-level markers. Word behaves the same way: it opened a redline with a deleted bookmark tracked directly in
the body and an inserted one tracked directly in a table cell without complaint, kept both through Reject All
Changes, and opened a three-reviewer consolidated redline of the same shape. Neither application loses a
bookmark Docxodus tracks this way; only Docxodus's own accept and reject resolve it per side.

### Docxodus code

- `RevisionProcessor.AcceptDeletedAndMoveFromParagraphMarksTransform` rebuilt each block stream from its
  paragraphs and tables only, dropping every other child. It now keeps them (`LooseChildrenOfBlockStream`) in
  front of the block they preceded; a marker between paragraphs that a deleted paragraph mark merges goes inside
  the merged paragraph, unless only deleted content precedes it.
- `RevisionProcessor.ReverseRevisionsTransform` reversed a `w:ins`/`w:del` only under a paragraph or run
  container; a block stream (`IsRunLevelRevisionParent`) now counts too.
- The IR still leaves the markers out (rule N3). `IrMarkupRenderer.PlanBlockLevelMarkers` and
  `PlanCompositeBlockLevelMarkers` put them back; see *Block-level markers* in
  `docs/architecture/ir_diff_engine.md`.

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

N-way consolidation follows the same policy (issue #1020). `DocxCompare.Consolidate`, and every
transport's consolidate op through it, accepts the base and every reviewer copy first. Word's Combine
would instead keep the inputs' revisions, but the consolidate renderer cannot carry them through: before
this policy the engine merged the accepted body while copying unchanged headers, footers and notes from the
base verbatim, so their pending markup leaked into the redline under its original author and rejecting all
changes no longer reproduced either the base or its accepted view in those parts.

### Relevant code

- `Docxodus/DocxCompare.cs` — `ApplyFrontDoorRevisionPolicy` (pre-accept only) and its N-way counterpart
  `ApplyFrontDoorConsolidatePolicy`, which `Consolidate` and `Internal/DocxDiffOps.cs` apply.
- `Docxodus/DocxDiff.cs` — `PreAcceptInputRevisions`, `PreserveInputRevisions`.

### Tests

`Docxodus.Tests/DocxCompareOriginalPendingRevisionsTests.cs` pins every row above through the front door
(redline markup, accept-all, reject-all, and that no change by the earlier author survives);
`DocxCompareTests.FrontDoorPolicy_*` pin the policy flags themselves.
`ConsolidateFrontDoorTests` pins the N-way policy through `DocxCompare.Consolidate`, the facade, the
stdio host and the MCP consolidate form.

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

## Text frames: a partial direct `w:framePr` combines with the style's

A paragraph style can define a text frame, and a paragraph that uses the style can carry its own
`w:framePr` with only some attributes. Here the paragraph sets only `w:wrap`:

```xml
<!-- styles.xml -->
<w:style w:type="paragraph" w:customStyle="1" w:styleId="StampFrame">
  <w:name w:val="Stamp Frame"/>
  <w:pPr><w:framePr w:w="8640" w:wrap="auto" w:vAnchor="text" w:hAnchor="page" w:x="1008" w:y="40"/></w:pPr>
</w:style>

<!-- footer1.xml -->
<w:p>
  <w:pPr><w:pStyle w:val="StampFrame"/><w:framePr w:wrap="auto"/></w:pPr>
  <w:r><w:t>REF 55-0912 rev3</w:t></w:r>
</w:p>
<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>Leaf 4</w:t></w:r></w:p>
```

### Comparison

Letter page, 1in margins. The x column is where the stamp text begins.

| Paragraph's own `w:framePr` | LibreOffice 25.8.7.3 (PDF export) | Docxodus before | Docxodus after |
|---|---|---|---|
| none (style only) | x = 72.1pt: no frame, the paragraph stays in the footer's flow | frame at `left: 50.4pt`, 432pt wide, page-anchored | unchanged |
| `w:wrap="auto"` | x = 50.5pt: the style's `w:x` of 1008 twips | frame at `left: 0`, no width, text-anchored | frame at `left: 50.4pt`, 432pt wide, page-anchored |
| `w:wrap="auto" w:x="2000"` | x = 100.1pt: the paragraph's `w:x`, the style's other attributes | frame at `left: 100pt`, no width, text-anchored | frame at `left: 100pt`, 432pt wide, page-anchored |

Word was not tested.

### Analysis

LibreOffice resolves a direct `w:framePr` attribute by attribute: each attribute the paragraph sets
wins, and each it omits comes from the style. Docxodus already resolves `w:spacing` and `w:ind` that
way, but `w:framePr` was on no list of properties to merge attribute by attribute, so the
paragraph's element replaced the style's whole and the frame lost its width, its horizontal anchor
and its offsets.

LibreOffice also ignores a frame that only the style defines (first row). Docxodus renders that frame,
and this fix leaves that case alone.

### Two exceptions: alignment pairs and drop caps

A plain attribute-by-attribute merge is wrong in two places. Each probe below is a body paragraph
in a style `F`, on a Letter page with 1in margins, measured in a LibreOffice 25.8.7.3 PDF export.
The frame is 2880 twips (144pt) wide where the style sets `w:w`.

**`w:x`/`w:xAlign` and `w:y`/`w:yAlign` are one setting per axis.** Within one element, the
paginated view gives `w:xAlign` precedence over `w:x` and `w:yAlign` over `w:y`. A naive merge
combines a style's `w:xAlign` with the paragraph's own `w:x` into a pair neither layer wrote, so the
paragraph's offset is lost. LibreOffice keeps the offset the paragraph set:

| Style `w:framePr` | Paragraph `w:framePr` | LibreOffice frame text |
|---|---|---|
| `w:w="2880" w:wrap="around" w:vAnchor="page" w:hAnchor="page" w:xAlign="center" w:y="4000"` | `w:wrap="around" w:x="2000"` | x = 100.1pt (the paragraph's `w:x`, not centred) |
| same | restates every attribute, with `w:x="2000"` in place of `w:xAlign` | x = 100.1pt |
| `w:w="2880" w:wrap="around" w:vAnchor="page" w:hAnchor="page" w:x="2000" w:yAlign="center"` | `w:wrap="around" w:y="4000"` | y = 200pt (the paragraph's `w:y`, not centred) |
| `w:w="2880" w:wrap="around" w:vAnchor="page" w:hAnchor="page" w:x="2000" w:y="4000"` | `w:wrap="around" w:xAlign="right"` | x = 468pt (right-aligned on the page) |
| `w:w="2880" w:wrap="around" w:vAnchor="page" w:hAnchor="margin" w:x="2000" w:y="4000"` | `w:wrap="around" w:hAnchor="page" w:xAlign="center"` | x = 234pt (centred on the page) |

The anchors are separate settings, not part of their axis. A paragraph that sets only `w:hAnchor`
keeps the style's offset or alignment and measures it from its own anchor:

| Style `w:framePr` | Paragraph `w:framePr` | LibreOffice frame text |
|---|---|---|
| `w:w="2880" w:wrap="around" w:vAnchor="page" w:hAnchor="margin" w:xAlign="right" w:y="4000"` | `w:wrap="around" w:hAnchor="page"` | x = 468pt (right-aligned on the page) |
| `w:w="2880" w:wrap="around" w:vAnchor="page" w:hAnchor="page" w:x="2000" w:y="4000"` | `w:wrap="around" w:hAnchor="margin"` | x = 172.1pt (the style's 100pt from the margin) |
| `w:w="2880" w:wrap="around" w:vAnchor="margin" w:hAnchor="page" w:x="2000" w:y="2000"` | `w:wrap="around" w:vAnchor="page"` | y = 100pt (the style's 100pt from the page) |

**A style's `w:dropCap` never reaches a paragraph with its own `w:framePr`.** The style below
defines a three-line drop cap. Converting each probe to ODT shows whether LibreOffice built a drop
cap (`style:drop-cap`) or an ordinary frame (`draw:frame`):

| Style `w:framePr` | Paragraph `w:framePr` | LibreOffice |
|---|---|---|
| `w:dropCap="drop" w:lines="3" w:wrap="around" w:vAnchor="text" w:hAnchor="text"` | `w:wrap="around" w:hAnchor="page" w:x="4000" w:vAnchor="page" w:y="4000"` | ordinary frame at (200pt, 200pt) |
| same | `w:wrap="around"`; `w:w="2880"`; `w:hAnchor="page"`; `w:x="4000"`; or `w:vAnchor="page" w:y="4000"` | ordinary frame in each case |
| same | `w:lines="2"` | ordinary frame |
| `w:dropCap="drop" w:lines="3" w:wrap="around" w:hAnchor="page" w:x="4000" w:vAnchor="page" w:y="4000"` | `w:wrap="around"` | ordinary frame at (200pt, 200pt) |
| none | `w:dropCap="drop" w:lines="3" w:wrap="around" w:hAnchor="page" w:x="4000" w:vAnchor="page" w:y="4000"` | drop cap, positioning ignored |
| `w:dropCap="drop" w:lines="3" w:wrap="around" w:vAnchor="text" w:hAnchor="text"` | none | no frame and no drop cap (the style-only rule above) |

Docxodus treats a paragraph whose merged frame has a `w:dropCap` other than `none` as a drop cap and
renders no frame for it. The last row shows LibreOffice also ignores a drop cap that only the style
defines. Docxodus keeps treating that paragraph as a drop cap, as it did before this fix. Whether a
`w:dropCap` passes from a parent style to a `w:basedOn` style that has its own `w:framePr` cannot be
measured this way, because LibreOffice ignores style-only frames. Docxodus applies the same rule
there and does not inherit it. Word was not tested.

### Code involved

`FormattingAssembler.MergeStyleElement` merges each pair of property layers (style chain, table style,
numbering and direct formatting). `w:framePr` is now one of its special-cased children and goes through
`FramePrMerge`. When the higher-priority layer sets `w:x` or `w:xAlign`, it drops both from the lower
layer, and likewise for `w:y` and `w:yAlign`. It always drops `w:dropCap` and `w:lines`. It then hands
the rest to `AttributeMerge`, the same higher-priority-attributes-first merge `w:spacing` uses.
`WmlToHtmlConverter.CreateBorderDivs` reads the merged element.

### Tests

`Docxodus.Tests/FramePrStyleMergeTests.cs`.

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
`normal` an explicit height for its own font, measured at 10000 px to reduce rounding of its
ascent, descent and line gap.
The converter's output is unchanged. Since issue #1151 the paginator itself runs the same pass on its
staging tree, so the paginated view and the paginated editor lay single-spaced lines out at the same
unrounded pitch as the export; after the export's earlier pass no line height is `normal`, so the
paginator's pass changes nothing there. The paginated editor runs it again on each block it renders
afresh after an edit (`wireBlock` in `npm/src/editor.ts`), so an edited paragraph keeps the page's
pitch. For the 11pt Times New Roman body of the NVCA model charter that is 12.65pt a line instead of
12.75pt, and 11.50pt instead of 12pt for its 10pt notes (Times New Roman's `hhea` height is
2355/2048 = 1.1499 em; LibreOffice advances 12.65pt and 11.55pt). Word's baseline placement
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
  Since issue #1151 the read-only paginated view (`paginateHtml`, the React viewer) runs it too, at
  the end of pagination (`PaginationOptions.alignBaselinesToWord`); the paginated editor does not,
  since text typed into a moved box would not move with it. Once the view's lines took the font's
  unrounded height (#1151), Chromium's whole-pixel baseline rounding fell differently and moved
  some first lines a pixel; on the seven corpus cases with a Word reference, the aligned view's ink
  extents match Word's or come within one pixel of them.
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
no runs keeps its mark's font.

Tests: `HCO100_*` in `Docxodus.Tests/HtmlConversionOpsTests.cs`, and the plain-mark case in
`npm/tests/export-line-baselines.spec.ts`, which serves Calibri and Arial as two different faces.

### A paragraph mark taller than its runs does not make any line taller (#949)

The pilcrow sits on a paragraph's last line, so a mark much larger than the text might be expected to make
that line taller. Word for the web says it does not. `line-baselines.docx` CASE8 and CASE9 are three lines
of 10 pt Calibri under a 20 pt Calibri mark, at single and 1.15 spacing. Each is followed on its page by a
plain 10 pt paragraph (AFTER8, AFTER9), whose first baseline shows how tall the last line was:

```xml
<w:p><w:pPr><w:spacing w:line="240" w:lineRule="auto"/>
    <w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="40"/></w:rPr></w:pPr>
  <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="20"/></w:rPr>
    <w:t>CASE8 The quick brown fox…</w:t></w:r></w:p>
```

| Baselines, pt from the page top | line 1 | line 2 | line 3 (holds the mark) | next paragraph |
|---|---:|---:|---:|---:|
| Word, single (CASE8) | 81.53 | 93.78 | 105.78 | 118.05 |
| Word, 1.15 (CASE9) | 81.53 | 95.53 | 109.53 | 123.55 |

Every step is one ordinary 10 pt Calibri line: 12.0–12.27 pt single, 14.0–14.02 pt at 1.15 (10 pt Calibri's
natural height is 12.21 pt). The first baseline, 81.53 pt, is 10 pt Calibri's ascent below the margin. The
20 pt mark grows neither the last line nor any other.

Docxodus already matches, because `DefineStrutFont` (#940) sizes the paragraph's strut from its runs and
the mark draws no glyph. The fixture case is a regression guard: with the strut at the mark's 20 pt, the
second baseline lands 15.96 pt below the first instead of Word's 12.25 pt.

Tests: the #949 case in `npm/tests/export-line-baselines.spec.ts`.

### Precision loss before auto-spacing (#1124)

The related #1124 controls in `TestFiles/BaselineFlow/word-observations.json` expose a smaller
precision loss in plain-paragraph export. For example, `w:spacing w:line="276"
w:lineRule="auto"` means a 1.15 multiple. Rounding the resolved natural height to twips and
flooring the product shortens every line. Native Word long-paragraph controls at 10, 13 and
16 points demonstrate the accumulated error independently of the paragraph's page origin.

| Measurement on the 13pt Arial control | Baseline pitch / accumulated difference |
| --- | --- |
| Native Word for the web PDF | 17.2014pt average within the captured page segment |
| Docxodus before the precision correction | 17.15pt pitch; up to 0.72pt relative baseline error |
| LibreOffice | Not measured for this specific control |

`npm/src/word-text.ts` now retains the resolved font height through the spacing multiple.
The shared probe in `npm/src/line-metrics.ts` also measures at 10000px: a 1000px Liberation
Serif probe reports a 1.149 height ratio, whereas 10000px reports 1.1499, close to the font's
2355/2048 ratio. The native assertions in `npm/tests/word-baselines.spec.ts` cover both this
probe loss and premature line rounding. These are relative-pitch observations, not a claim
that every Word PDF glyph or raster matches.

Matching line inventories also leave a distinct horizontal error when the painter copies
the browser's hinted word starts. Native 11pt and 12pt glyph-gap controls retain each raw PDF
coordinate; independent font-table expectations check natural word advances and alignment
without assuming that Word's PDF placement quantization is a universal grid. The zero-fit
path now preserves inline flow geometry while painting the same unrounded font advances
used for newly fitted lines. Unsupported break controls retain their existing DOM path.

The original agreement exposes a provider conflict: with the exact Liberation Serif glyph
outlines and advances embedded in the Word capture, retaining the auto-spacing precision
improves the local Word raster diagnostic while reducing its agreement with the preserved
LibreOffice reference. Both raw references and the historical score remain unchanged. This
observation does not exempt the case from the existing CI regression gate.

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

## Tables: where the first cell's text starts (issue #1219)

Three rules decide how far right of the indent a table's first cell text starts. They were measured in
PDFs exported by Word for the web (File > Export > Download as PDF, October 23, 2026; no build reported)
from authored one-cell tables, US Letter with 72pt margins and 1pt (`w:sz="8"`) borders:

1. **A cell margin no one states is zero.** With no `w:tcMar`, no `w:tblCellMar` and no table style,
   the text sits just inside the border. The 108-twip (5.4pt) margin a new Word table shows comes from
   the document's default table style ("Normal Table", `w:default="1"`), which also formats every
   table that names no style.
2. **Mode 15 puts the table at the indent.** In compatibility mode 15 the leading border's outer edge
   is at the indent (`w:tblInd`, zero when absent), and the text follows at the cell margin, counted
   from the border's centre line.
3. **Before mode 15 the text goes to the indent.** In modes 14 and 12, and when the settings state no
   mode at all, Word pulls a top-level, start-aligned table out by its first cell's own left margin, so
   that cell's text starts exactly at the indent. It does this whenever the table states `w:tblInd`, or
   the table style governing it (its own or the default one) states an indent or cell margins. A table
   with none of these has its leading border centred on the indent. A centred table and a nested
   table are placed as in mode 15, and a first row that starts with `w:gridBefore`, or a 100%-wide
   table, is still pulled.

```xml
<w:settings><w:compat><w:compatSetting w:name="compatibilityMode"
  w:uri="http://schemas.microsoft.com/office/word" w:val="14"/></w:compat></w:settings>
<w:tblPr><w:tblW w:w="5000" w:type="dxa"/><w:tblInd w:w="0" w:type="dxa"/>
  <w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr>
```

First-cell text start, in points right of the body text edge (Word's PDF positions sit on a 0.5pt
grid). Docxodus is the HTML converter's output laid out by Chromium, the engine behind the paginated
view and the PDF export:

| Table | Word, mode 15 | Word, mode 14/12/none | Docxodus before | Docxodus after, 15 | Docxodus after, 14/12/none |
|---|---:|---:|---:|---:|---:|
| no indent, 108-twip margins | 6.0 | 5.5 | 6.1 | 6.1 | 5.7 |
| `tblInd` 0, 108 | 6.0 | 0.0 | 6.1 | 6.1 | 0.3 |
| `tblInd` 720, 108 | 42.0 | 36.0 | 42.1 | 42.1 | 36.2 |
| `tblInd` -720, 108 | -30.0 | -36.0 | -29.9 | -29.9 | -35.8 |
| no indent, margins 0 | 1.0 | 0.5 | 0.8 | 0.8 | 0.3 |
| no indent, no margins, no table style | 1.0 | 0.5 | 6.1 | 0.8 | 0.3 |
| `tblInd` 0, first cell `tcMar` 288 | 15.0 | 0.0 | 15.1 | 15.1 | 0.3 |
| table style with `tblInd` 0 and 108 | 6.0 | 0.0 | 6.1 | 6.1 | 0.3 |
| names no style; default style has 108 and `tblInd` 0 | 6.0 | 0.0 | 6.1 | 6.1 | 0.3 |
| names no style; default style has 108 only | — | 0.0 | 6.1 | — | 0.3 |
| 3pt border, no indent, 108 | 7.0 | 5.5 | 8.4 | 8.4 | 6.9 |
| 3pt border, `tblInd` 0, margins 0 | 3.0 | 0.0 | 3.0 | 3.0 | 1.5 |
| named style with no indent and no margins | 1.0 | 0.5 | 6.1 | 0.8 | 0.3 |
| default style; first row `w:gridBefore` 1000 twips | — | 0.0 | — | — | 0.3 |
| default style; `w:tblW` 5000 pct (100%) | — | 0.0 | — | — | 0.3 |
| default style; nested table, from the outer cell's text | — | 6.0 | — | — | 6.4 |
| default style; `w:jc="center"` 250pt table, from the margin | — | 114.6 | — | — | (centred, not pulled) |

The PDFs' vector borders confirm rule 2's geometry: with a 3pt border the mode 15 table's left border
spans 72.0–75.0pt (outer edge at the margin) and the mode 14 table's 70.5–73.5pt (centred on it).

**Docxodus.** `ResolveTableIndentInputs` runs before formatting assembly, while `w:tblStyle` is still
there. A table that names no style, or names one the document does not define, gets the default table
style's `w:tblInd` and side cell margins wherever it states none. Before mode 15 the pass also records
whether the table is pulled out to its text or only centres its border. It does this only for
top-level tables aligned at the start; end-aligned tables are left alone too, though Word's handling of
them is unmeasured. A document with no settings part at all is treated like one whose settings state
no mode, which is also unmeasured.

The rows from "named style" down were captured on October 23, 2026 from six further fixtures
(`i1219b-*`). In the centred row, a pulled table would put its text at about 109pt. `CellMarginTwips` is the one margin lookup (cell, then table, then zero); the
cell padding and the tab-stop wrap width both use it. `ProcessTable` moves the table edge out by the
first cell's margin and half its leading border, in the table's leading direction; a floating table
(`w:tblpPr`) is placed by its own properties and is left alone. The 5.4pt padding the converter used to
give every cell is gone.

**Remaining difference.** Word measures a cell margin from the border's centre line and never lets text
start under the border; CSS pads from the border's inner edge. With a 1pt border that leaves Docxodus's
text about 0.25pt right of Word's in every mode; with a 3pt border, about 1.5pt.

- Code: `Docxodus/WmlToHtmlConverter.cs` (`ResolveTableIndentInputs`, `CellMarginTwips`,
  `ProcessTable`).
- Tests: `TableHorizontalPlacementTests`, `npm/tests/export-table-indent.spec.ts`; bench cases
  `tables-cell-margin-*`, `tables-indent-mode-*` and `tables-legacy-*` (core-v65).

The October 25 review re-exported the exact six legacy benchmark inputs through Word and retained
complete token inventories alongside their horizontal marker origins. A `gridBefore` control exposed
an additional width error: the HTML column group included a leading column omitted by every row,
so CSS applied that column's width to the first actual cell and split `BravoCell` and `gridBefore.`.
The column group now skips only a grid prefix omitted by all of the table's own rows. Nested rows
use their own grid, rows wrapped in content controls still count, and a column used by another row
is retained. The first-cell indent lookup also follows cells inside content controls. The correction checks
resolved alignment after formatting assembly, so a centred or right-aligned table inherits the
same placement whether alignment is stated directly or through its table style.

The original nested CI input has two SDK property-order findings and requests Liberation Serif,
which Word substitutes. Its untouched native capture remains an ineligible diagnostic. Two SDK-valid
controls, one preserving its requested font and one declaring Times New Roman, confirm the same
left-cell origins (73.025pt outer, 73.775pt inner). The old renderer adds 5.4pt at each level.
Whole-table automatic column allocation and the existing vertical offsets are separately retained
in the benchmark evidence; matching these horizontal margins does not establish a full Word score.

## `w:rtl` and `w:bidiVisual` are on/off values, and Google Docs writes them switched off

Both are `ST_OnOff` properties. Google Docs' `.docx` export writes the run-level one, switched off, on almost
every run (issue #1011):

```xml
<w:p>
  <w:r>
    <w:rPr><w:rtl w:val="0"/></w:rPr>
    <w:t>2026. 10. 07.</w:t>
  </w:r>
</w:p>
```

| Renderer | Result |
|---|---|
| Word | Left-to-right run: `2026. 10. 07.` |
| Docxodus (before) | Run wrapped in RLM marks (U+200F); the browser displays roughly `.07 .10 .2026` |
| Docxodus (after) | No directional marks, the same as with no `w:rtl` |

The converter tested only whether `w:rtl` (and, for tables, `w:bidiVisual`) was present. `w:val` values `0`,
`false` and `off` switch the property off; a missing `w:val` means on. Strongly left-to-right text such as
Hangul hid the defect in most runs. It showed wherever a run started or ended with digits or punctuation,
which the bidi algorithm placed by the surrounding RLM marks.

- Code: `Docxodus/WmlToHtmlConverter.cs`: `DetermineRunMarks`, `ProcessTable`,
  `ReverseTableBordersForRtlTables`, all reading through `IsOnOffPropertyEnabled`.
- Tests: `Docxodus.Tests/WmlToHtmlOnOffDirectionTests.cs`.

## A right tab stop inside the right indent: the page number reaches the stop

Table-of-contents and index entries often set a right indent, so long titles wrap short of the
page-number column, while the right leader tab sits at the column edge (issue #1275):

```xml
<w:p>
  <w:pPr>
    <w:tabs><w:tab w:val="left" w:pos="1296"/><w:tab w:val="right" w:leader="dot" w:pos="9360"/></w:tabs>
    <w:ind w:left="1296" w:right="864" w:hanging="648"/>
  </w:pPr>
  <w:r><w:t>2.1.</w:t></w:r><w:r><w:tab/></w:r>
  <w:r><w:t>Quarry Access Roads</w:t></w:r>
  <w:r><w:tab/><w:t>3</w:t></w:r>
</w:p>
```

On a Letter page with one-inch margins the column is 9360 twips, so the line ends at 8496 twips but
the stop is at 9360. The same entry with an empty `w:bookmarkStart`/`w:bookmarkEnd` pair, a
`w:noBreakHyphen` run after the title, or `w:u`/`w:strike` on the tab runs should lay out the same way.

| Renderer | Result |
|---|---|
| LibreOffice 25.8.7.3 (PDF export) | Every entry on one line, page number ending at 540pt (the stop) in all five variants, with and without `w:right` |
| Word | Not measured |
| Docxodus (before) | With `w:right`: every entry on two lines. With a bookmark, non-breaking hyphen, underline or strike: the right tab used a fixed server-measured width, so the number ended short of the stop |
| Docxodus (after) | Every entry on one line; page number ends on the stop (624px in a 6.5in column) |

The stop can also lie past the column edge. LibreOffice keeps the entry on one line and ends the
number on the stop, in the page margin. With the stop at 10000 twips, the number ends at 572.1pt
(72 + 500pt), with or without `w:right="864"`; at 9800 it ends at 562.1pt. Docxodus before the fix
put these entries on two lines, with or without the right indent. After it, they stay on one line
and the number ends at 666.7px, the stop. A stop past the paper's edge (12000 twips, 672pt on a
612pt page) was not compared: LibreOffice runs the leader to the paper edge and the number does not
appear in its PDF.

Two separate causes:

- The right-tab segment is an `inline-flex` box whose width runs from its start to the stop, so it
  is 864 twips wider than what is left of the line. The browser moved it onto its own line. The tab
  pass now records how far the segment passes the line end (`pt:TabOverhang`): into the right indent,
  and past the column edge when the stop is there. The segment gets a matching negative
  `margin-right`, which lets it fit on the line while it keeps its full width.
- `ConvertAlignedTabRuns` accepted only runs that contain `w:rPr`, `w:t` and `w:tab`, and
  `CanFlexTabRun` rejected underline and strike. Every other paragraph fell back to the fixed-width
  tab. Bookmarks render as empty anchors and a non-breaking hyphen as text, so both now stay on the
  flex path. A bookmark that opens a rendered annotation range still falls back, as an annotation
  range that is already open does. Underline and strike were rejected, but CSS text decoration
  does not paint an empty tab box on either path, so the rejection cost alignment and bought no
  decoration. LibreOffice does draw the underline (and the strike) across an underlined (or struck)
  tab. Since issue #1285 the tab box draws them too, as borders of empty blocks inside it.
- A bookmark inside a `w:hyperlink` rendered as an `<a id>` inside the link's `<a href>`. HTML
  forbids nested links, so the browser's parser closed the link at the bookmark and left the label,
  title and number outside it, off their stops (issue #1285). Inside a link the bookmark target is
  now an empty `<span id>`.

- Code: `Docxodus/WmlToHtmlConverter.cs`: `CalculateSpanWidthTransform` (right-tab branch),
  `ConvertAlignedTabRuns`, `CanFlexTabRun`.
- Tests: `Docxodus.Tests/RightTabLeaderLineTests.cs`, `npm/tests/right-tab-leader-lines.spec.ts`.

## Leader lines with field page numbers, wrapping titles, and a second stop in the indent (issue #1285)

Three more shapes of the entry above (left stop at 1296, `w:ind w:left="1296" w:right="864"
w:hanging="648"`, Letter page, one-inch margins, so the line ends at 8496 twips):

- The page number is a `PAGEREF` complex field (`w:fldChar` begin, `w:instrText`, separate, result,
  end), as in every Word-built table of contents.
- The title runs past the line end and wraps, so the page number belongs on its last line.
- A first right stop at 8800 lies inside the right indent, and a second stop follows it: a right stop
  at 9360 or a left stop at 8900.

| Shape | LibreOffice 25.8.7.3 | Docxodus (before) | Docxodus (after) |
|---|---|---|---|
| `PAGEREF` number | One line, number ending on the stop (540pt) | Fixed server-estimated tab: number on the next line (with `w:right`), or ending at 609.5px and 617px instead of 624px (without) | One line, number ending at 624px |
| Wrapping title | Title on two lines, number at the end of line 2 on the stop (540.1pt), leader from the title's last word | Number on a third line, at 528–544px | Title on two lines, number at the end of line 2 ending at 624px, leader from the last word |
| Right 8800, then right 9360 | One line: "12" ends at 512pt (8800), "34" at 540.1pt (9360) | Both numbers on line 2, at 86–160px | One line: "12" ends at 586.7px, "34" at 624px |
| Right 8800, then left 8900 | One line: "A" starts at 517.1pt (8900) | Line 2 | One line: "A" starts at 593.3px |
| Word | Not measured | | |

LibreOffice keeps text after the left stop at 8900 on the line only while it fits before the column
edge (9360); a wider word moves to the next line, behind a leader. Forcing a 24-word suffix into the
nowrap tail extended it beyond the page; only seven words fit within the column. The tab-width pass
now identifies overflowing text after that final left tab and keeps the paragraph on its ordinary
wrapping path. All 24 words then stay in the column. The preceding number retains the ordinary path's
different placement from LibreOffice, visible in the benchmark captures. Word is unmeasured.

Causes and fixes, all in `Docxodus/WmlToHtmlConverter.cs`:

- `ConvertAlignedTabRuns` refused every run that held a `w:fldChar` or `w:instrText`, so every field
  paragraph used the fixed tab widths. Those runs draw nothing, so they now stay on the flex path. A
  stretch converts its fields on its own, so a field open across a tab is cut in two; a `HYPERLINK`
  (rendered as one link) or a `PAGE`/`NUMPAGES` marker cut that way keeps the old path, as does a form
  field, whose begin carries form data.
- `CalculateSpanWidthTransform` marked any tab whose pen was past the line end as following wrapped
  text (issue #891). After a right stretch ends on a stop inside the indent, the pen is past the line
  end without any wrap; the pass now allows the pen to reach the furthest such stop. In
  `ConvertAlignedTabRuns`, the segment that ends there hands its overhang back with a negative
  `margin-right`, which would also pull everything after it left by that much. So what follows goes in
  one nowrap, content-wide box whose left margin undoes the overhang and whose `-100%` right margin
  takes it off the line (a zero-width box let a `pre-wrap` run wrap inside it in the paginated view).
- A right tab after wrapped text kept the fixed path, sized from an estimate of where the last line
  ends. When the wrapped text began at the left indent, where its later lines also begin, the pass
  records the stop's distance from that indent (`pt:TabWrappedSegmentWidth`). The segment then spans
  from the indent to the stop: the text wraps inside it at the line end, `align-items: last baseline`
  puts the number on the text's last line, and the tab, which takes no room after the last word,
  draws a leader as wide as the segment, clipped (`overflow-x: clip`) where the number begins.
  Wrapped text that began elsewhere keeps the fixed path.

- Tests: `Docxodus.Tests/RightTabLeaderLineTests.cs`, `npm/tests/right-tab-leader-lines.spec.ts`.

## `w:defaultTabStop` of 0 or a few twips: automatic stops past the pen, and none at 0 (issue #1269)

Automatic tab stops are the interval multiples that follow the paragraph's last custom stop.
ECMA-376 Part 1 §17.15.1.25 calls `w:defaultTabStop` "the multiplier to generate automatic tab
stops", says that if it is omitted "automatic tab stops should be generated at 720 twentieths of a
point (0.5") intervals across the displayed page", and types `w:val` as `ST_TwipsMeasure`, so 0 is
schema-valid. Microsoft's implementer notes (MS-OI29500 revision 25.0, 18 August 2026, §2.1.399)
add only: "Word defines the values of this attribute with the XML Schema short datatype, with a
minimum value of 0." Neither document says how a zero interval lays out.

```xml
<!-- settings.xml: 12, or 0 -->
<w:defaultTabStop w:val="12"/>
<!-- document.xml: no w:tabs, no indent -->
<w:p>
  <w:r><w:t>Observatory inventory</w:t></w:r><w:r><w:tab/></w:r>
  <w:r><w:t>North</w:t></w:r><w:r><w:tab/></w:r>
  <w:r><w:t>East</w:t></w:r><w:r><w:tab/></w:r>
  <w:r><w:t>South</w:t></w:r><w:r><w:t>Copper lens</w:t></w:r>
</w:p>
```

Positions are each word's left edge on the PDF page (left margin at 72pt), from the issue's three
inputs. "inventory" ends at 179.99pt in Word and 180.34pt in LibreOffice.

| Renderer | `w:val="12"` | `w:val="0"` | `w:val="720"` (control) |
|---|---|---|---|
| Word for the web (Download as PDF; Times New Roman substituted for the requested font, metric-compatible) | North 180.58, East 208.85, South 229.85: each tab advances to the next 12-twip multiple, under 0.6pt | North 180.33, East 208.33, South 229.10: each tab advances nothing (0.34/0.02/0.20pt residue, within Word's positioning grid) | North 216.10, East 252.10, South 288.13: the half-inch grid |
| LibreOffice 25.8.7.3 (PDF export) | Each tab advances to the next 12-twip multiple, under 0.6pt | North 213.9, East 249.4, South 284.8: a 1.25cm grid, LibreOffice's own default interval | North 216.1, East 252.1, South 288.1 |
| Docxodus (before) | `InvalidOperationException: Sequence contains no matching element` | Same exception | Half-inch grid |
| Docxodus (after) | Each tab advances to the next 12-twip multiple (0.003in–0.007in) | Each tab advances nothing | Unchanged |

Word therefore reads 0 literally: there are no automatic stops, so a tab past the last custom stop
has no width. It does not fall back to the half-inch default, as it would if the element were
omitted, and it does not impose a minimum interval at 12 twips. LibreOffice differs at 0.

Why it threw: the tab pass generated only 100 automatic stops up front, which is 1,200 twips at a
12-twip interval and nothing at all at 0 (every stop sat at position 0). Once the pen passed the last
generated stop, the pass gave that tab a half-inch width and stopped measuring the paragraph, so the
later tabs had no width, rendered no tab element, and the flexible-tab renderer failed looking for one.

The fix: past the stops generated up front, the next automatic stop is computed on demand as the first
interval multiple beyond the pen, so every tab is measured and the pass never stops early. A zero
interval generates no automatic stops, as in Word: a tab with no stop ahead of it advances nothing.
Custom stops (and the left-indent stop) still apply, and a tab past them also advances nothing.
That part is inferred, not measured: the Word captures above have no custom stops or indent. A
negative value is schema-invalid; it is read as 0, Word's minimum, rather than as an absent element.
The next automatic stop after a position is the floor of position / interval, plus one, times the
interval, so a pen left of the margin (a negative indent) does not skip the multiple just after it.
This rounding is also inferred: it matches LibreOffice and the spec's definition of automatic stops
as interval multiples, but was not measured in Word.

What remains: the tab pass chooses each stop from its own estimate of the preceding text's width. In
the browser export, where no font metrics are available to that pass, the estimate ran about 10pt
wide on the lead text. At 12 twips each word therefore sits several points right of where LibreOffice
puts it (`TabNorth` at 237.6pt against 226.9pt in the fidelity bench's `tabs-default-interval-12`).
A zero-width tab shows the same offset, since the text before it is laid out at its estimated width:
in the issue's `tabs-0.docx` the browser export puts North at 185.0pt, where Word puts it at 180.3pt.
A half-inch grid absorbs the error, so the 720 control matches LibreOffice to 0.1pt.

- Code: `Docxodus/WmlToHtmlConverter.cs`: `CalculateSpanWidthForTabs` (interval), `CalculateSpanWidthTransform` (stop lookup).
- Tests: `Docxodus.Tests/DefaultTabStopIntervalTests.cs`.

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

## Footnote numbers are the note's own reference run, and Word raises superscripts less than the browser (issue #1003)

**Reproducer** (`TestFiles/CA/CA008-Footnote-Reference.docx`, `word/footnotes.xml`):

```xml
<w:footnote w:id="1">
  <w:p>
    <w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>            <!-- 10 pt -->
    <w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>  <!-- vertAlign superscript -->
      <w:footnoteRef/></w:r>
    <w:r><w:t xml:space="preserve"> This is a test.</w:t></w:r>
  </w:p>
</w:footnote>
```

| Renderer | Number baseline above the note text's baseline |
|---|---|
| Word | 3 pt (0.30 em) |
| LibreOffice | not measured |
| Docxodus before #1003 (generic number in the page font, CSS `super` at 0.85 em) | 4.66 pt |
| Docxodus after #1003 (the note's own run, `<sup>` at the note's size) | 4.59 pt |
| Docxodus after #1016 (Word's superscript size and raise, below) | 3.45 pt |

**Analysis.** Word draws the number by rendering the `w:footnoteRef` run with that run's own formatting, so the number takes the FootnoteReference style's superscript at the note text's size. The paginated converter now does the same: `RenderPaginatedFootnoteNumber` converts a copy of that run with the display number in place of the marker. The rest of the gap came from how every superscript was raised, resolved by #1016 (next section).

**Code.** `WmlToHtmlConverter.RenderPaginatedFootnoteNumber` and the `.footnote-number[data-note-run]` rule.

## Superscript and subscript size and raise (issue #1016)

**Reproducer.** A run with `<w:vertAlign w:val="superscript"/>` (or `subscript`) next to an ordinary run of the same `w:sz`, in several faces and sizes.

Word reference PDFs of 8 faces at 8 to 288 pt give the following. Baselines in those PDFs fall on a 0.5 pt grid, so the large sizes are the precise ones:

| Face | Superscript raise (em) | Size | Subscript drop (em) | OS/2 `ySuperscriptYOffset` (em) |
|---|---|---|---|---|
| Calibri | 0.335 | 0.65 | ~0.095 | 0.477 |
| Aptos | 0.350 | 0.60 | ~0.075 | 0.350 |
| Times New Roman | 0.354 | 0.65 | ~0.076 | 0.453 |
| Arial | ~0.35 | 0.65 | ~0.073 | 0.477 |
| Verdana | 0.354 | 0.65 | ~0.073 | 0.453 |
| Georgia | 0.264 | 0.65 | ~0.076 | 0.265 |
| Courier New | 0.292 | 0.65 | ~0.105 | 0.422 |
| Cambria, up to 36 pt | 0.250 | 0.65 | ~0.08 | 0.246 |
| Cambria, 72 pt and up | 0.380 | 0.60 | ~0.088 | 0.246 |

| Renderer | Superscript | Subscript |
|---|---|---|
| Word | per face, above | per face, above |
| Chromium `vertical-align: super` / `sub` with `font-size: smaller` (Docxodus before) | 0.35 em (36 pt) to 0.41 em (8 pt), at 0.83 | 0.22 em (36 pt) to 0.27 em (8 pt), at 0.83 |
| Docxodus after | 0.345 em at 0.65 | 0.085 em at 0.65 |

**Analysis.** No single rule reproduces Word. Not a fixed fraction: the raise ranges from 0.25 to 0.38 em. Not the face's OS/2 superscript offset: Georgia and Aptos follow it, but Calibri, Arial, Times New Roman, Verdana and Courier New are raised well below it. And not one size-dependent rule: Cambria changes both its raise and its size between 36 and 72 pt. Word's size follows each face's OS/2 `ySuperscriptYSize` (0.65, Aptos 0.60) apart from Cambria's large sizes. HTML cannot express a per-face rule without per-face metrics, so the converter uses fixed values. These have the smallest worst-case error across the faces most documents use (Calibri, Aptos, Times New Roman, Arial): within 0.01 em of Word's superscript raise for all four. Georgia, Courier New and Cambria stay off by up to 0.1 em.

**Code.** `WmlToHtmlConverter.ScriptElement` and its constants, used by `ConvertRun` for `w:vertAlign`, by the footnote and endnote reference markers, and by the generic `.footnote-number` rule. `vertical-align` lengths are in the shrunken element's own em, so the raise is written as 0.345 / 0.65 em. `HtmlToWmlConverter` takes a `sup`/`sub` element's size from the text around it, so the smaller size does not round-trip into `w:sz`.

**Tests.** `WmlToHtmlScriptPositionTests`.

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

## Exact line spacing fixes the baseline independently of font size (#882)

An exact line can be shorter than the font it contains. Word keeps the declared line pitch and
places the baseline at approximately **80% of that height**, rather than centering the font box
or subtracting the font's descent. Glyphs can overlap adjacent lines and extend into the margins;
the measured Word PDFs clip them at the paper edge, not at every line or at the body margin.

The reproducer is a Letter page with 1in margins and three paragraphs of this form:

```xml
<w:p>
  <w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/></w:pPr>
  <w:r><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/>
    <w:sz w:val="32"/></w:rPr><w:t>Line 1 of the body, typography and spacing.</w:t></w:r>
</w:p>
```

First baseline below the top margin, in points:

| Exact line | Font size | Word | LibreOffice 25.8 | Docxodus before | Docxodus after |
|---:|---:|---:|---:|---:|---:|
| 12 | 10 | 9.53 | 9.60 | 9.00 | 9.60 |
| 12 | 16 | 9.53 | 9.60 | 11.25 | 9.60 |
| 12 | 24 | 9.53 | 9.60 | 14.25 | 9.60 |
| 18 | 16 | 14.53 | 14.40 | 14.25 | 14.40 |
| 24 | 16 | 19.28 | 19.20 | 17.25 | 19.20 |

Word's numbers come from the PDF text origins captured on 2026-10-10. Its PDFs identify the creator
as Microsoft Word and embed TimesNewRomanPSMT for the requested Liberation Serif. The browser
regression uses a renamed Liberation Serif subset with the same vertical metrics. LibreOffice's
numbers are the earlier measurements recorded in #882. The small deviations from `0.8 × line`
reflect Word's quantized text origins.

The full-page fixture has 54 paragraphs at 16pt on exact 12pt lines. Word keeps all 54 on one page,
with first and last baselines at 81.53pt and 717.725pt from the paper top. The initial issue assumed
Word kept descenders inside the line and clipped glyph tops; the captured output disproves that
assumption. Docxodus previously centered the glyph boxes and treated their overhang as misplaced
content, so the export's corrective pass moved line 54 onto a second page.

`WmlToHtmlConverter` now retains the exact-spacing distinction in the paragraph CSS and marks its
runs. `alignExactLineBaselines` in `npm/src/line-metrics.ts` measures Chromium's baseline for each
font shape and applies a relative offset after pagination has split paragraphs. Keeping those
offsets out of measurement also preserves the range fragmenter's exclusion of positioned content.
Line boxes retain their exact heights, including mixed font sizes and hyperlink runs.

`preserveExactLineInk` expands the body's paint clip only for inline text in a paragraph whose
line boxes fit the body. The permitted ink must remain inside the paper and outside header,
footer, and footnote stories. Misplaced blocks and other overflow still use the corrective pass.
The clip is an explicit pixel inset: Chromium ignores `overflow-clip-margin` when only one axis
clips. PageMap measurement and offline re-verification honor that same inset.

Fixtures and the numbers-only Word record: `npm/tests/fixtures/exact-line-spacing/`.
Tests: `npm/tests/exact-line-spacing.spec.ts`. The #848 corrective-pass test in
`standalone-export.spec.ts` now uses minimum-spaced text, keeping its overflow coverage independent
of the authored exact-spacing exception.

## Tracked Drawings

### LibreOffice fails on tracked anchored header groups with text boxes

**Status:** Renderer limitation reproduced; explicit failure handling added (#1045).
**Measured:** 2026-10-10, Linux, LibreOffice 24.2.7.2 and 25.8.7.3.

The smallest failing combination in this investigation is a `w:ins` or `w:del` around a run
containing an anchored DrawingML group in a header, where the group contains text boxes. The
revision wrapper has this shape (drawing internals abbreviated):

```xml
<w:hdr>
  <w:p>
    <w:ins w:id="7" w:author="Reviewer" w:date="2024-01-01T00:00:00Z">
      <w:r><w:drawing>
        <wp:anchor ...>
          <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">
            <wpg:wgp>...<wps:wsp>...<wps:txbx><w:txbxContent>...</w:txbxContent></wps:txbx>...</wps:wsp>...</wpg:wgp>
          </a:graphicData></a:graphic>
        </wp:anchor>
      </w:drawing></w:r>
    </w:ins>
  </w:p>
</w:hdr>
```

`tools/diffharness/lo/tracked_drawings_fixture.py` supplies the complete literal XML, namespace
declarations and relationships. It creates two rectangles with invented text, without images or
external content. Replacing `w:ins` with `w:del` reproduces the same import failure. Both directly
authored revisions and public `DocxCompare.Compare` output fail, so the comparison engine alone
does not explain the problem.

The isolation matrix varies four dimensions independently and includes plain, insertion and
deletion versions of each combination: 2 × 2 × 2 × 2 × 3 = 48 packages.

| Control, relative to the failing tracked header group | LibreOffice 24.2.7.2 | LibreOffice 25.8.7.3 |
|---|---|---|
| Anchored group with text boxes in a header, inserted or deleted | Application error, no PDF | Application error, no PDF |
| Same group without revision wrapper | Renders | Renders |
| Single shape with text box | Renders | Renders |
| Group without text boxes | Renders | Renders |
| Inline group with text boxes | Renders | Renders |
| Anchored group with text boxes in the body | Renders | Renders |
| VML header text-box controls, plain and tracked | Renders | Renders |
| Accepted/rejected DrawingML comparison views | Renders | Renders |

Each version rendered 46 of the 48 matrix packages. With the original controls, four generated
insertion/deletion comparisons, their accept/reject views and paragraph-mark diagnostics, each
version rendered 67 of 73 packages. The same six packages failed with `Unspecified Application
Error` and no fresh PDF: the two matrix revisions, two directly authored header-group revisions,
and two DrawingML comparison outputs. Exit codes can differ between launches; failure handling
uses the actual process result and artifact presence rather than requiring a particular code.

All 57 input packages and 16 generated views have zero Office 2019 Open XML SDK validator errors.
The runner ties every validation record to the fixture's SHA-256. Native drawing/picture counts
in accepted comparison views match the right input, and rejected views match the left input.
Word desktop was not measured in this investigation. Docxodus's comparison and revision processor
produce schema-valid packages with those content results; this finding concerns LibreOffice's
import/conversion path, rather than Docxodus's HTML renderer.

#### Alternative encoding and preservation policy

Tracking only the paragraph mark (`w:pPr/w:rPr/w:ins` or `w:del`), leaving the drawing run outside
the revision, allows conversion. It fails the content contract: both accepted and rejected views
retain the drawing. For an insertion, rejection must remove it; for a deletion, acceptance must
remove it. Those diagnostic fixtures and both views remain in the matrix.

Inline placement also converts, but changes positioning semantics. The VML control demonstrates
a supported import path; it does not establish a general lossless conversion of a DrawingML
group. No equivalent safe revision encoding was identified by these experiments. The core
comparison output therefore retains the native DrawingML and its revision wrappers.

`tools/diffharness/lo/lo_render.py` provides explicit conversion handling for Linux harness callers:

- It converts a byte copy using an isolated LibreOffice profile and temporary output directory.
- JSON reports the renderer version, process result, diagnostics, source SHA-256, revision/drawing
  metadata and source preservation. `status` is `rendered`, `failed` or `unavailable`.
- Nonzero exit, timeout, missing fresh PDF, invalid PDF signature and workspace/output failures
  are distinct failure codes. The CLI exits nonzero for every unsuccessful conversion.
- A timeout terminates the renderer's isolated process group. A successful PDF is published by
  replacing its destination entry, so an existing symlink or hard link cannot overwrite the DOCX.
- It preserves the input and performs no revision flattening or shape rewrite. A known drawing
  combination is diagnostic metadata, not a blanket rejection of unmeasured renderer versions.

`rendered` confirms a fresh PDF conversion and signature, not visual fidelity. Callers must honor
the JSON status and `pdf` field; a pre-existing destination PDF is not evidence of success.

#### Reproduce and verify

From the repository root, with .NET 10 and LibreOffice installed:

```sh
python3 tools/diffharness/lo/tracked_drawings_fixture.py "$PWD" /tmp/tracked-drawings-fixtures
python3 tools/diffharness/lo/tracked_drawings_matrix.py /tmp/tracked-drawings-fixtures /tmp/tracked-drawings-current
python3 tools/diffharness/lo/tracked_drawings_matrix.py /tmp/tracked-drawings-fixtures /tmp/tracked-drawings-older --soffice /path/to/older/soffice
python3 -m unittest discover -s tools/diffharness/lo -p 'test_*.py'
```

Use an empty fixture directory initially. Each matrix run writes `results.json` containing every
conversion observation, including the six failures. Exit 0 means the measured baseline matches
24.2.7.2 or 25.8.7.3; it does not mean every fixture rendered. Unexpected control failures or
changed measured behavior exit 1. An unmeasured version exits 2 unless a control failed, and its
observations remain available for investigation. The older version was the official 24.2.7.2
Linux archive, extracted locally without changing the system installation.

Relevant core paths are `DocxCompare.Compare`, `IrMarkupRenderer`'s header markup and
`RevisionProcessor.AcceptRevisions` / `RejectRevisions`. This change adds harness handling and
interoperability evidence without changing those core paths. Unit tests cover failure reporting,
artifact freshness, process cleanup and source preservation, and run in CI without requiring a
LibreOffice installation.


## Fields: a `PAGE` result is formatted by the field, not by its cached result run (issue #1115)

**Status:** Documented, Docxodus matches the reference (October 2026)

A page-number field's cached result can carry run formatting that the field's instruction does not:

```xml
<w:fldSimple w:instr=" PAGE \* Arabic "><w:r><w:rPr><w:i/></w:rPr><w:t>1</w:t></w:r></w:fldSimple>
<!-- complex form: the instruction run is unformatted, only the result run is italic -->
<w:r><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:instrText xml:space="preserve"> PAGE \* Arabic </w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="separate"/></w:r>
<w:r><w:rPr><w:i/></w:rPr><w:t>1</w:t></w:r>
<w:r><w:fldChar w:fldCharType="end"/></w:r>
```

| Form | Word | LibreOffice 25.8.7.3 | Docxodus (paginated) |
|---|---|---|---|
| Simple `PAGE` with italic cached result | Not yet captured | Upright current page number | Upright current page number (from #1115) |
| Complex `PAGE`, unformatted instruction, italic result | Not yet captured | Upright current page number | Upright current page number |
| Legacy `w:pgNum` in an italic run | Not yet captured | Block dropped | Italic decimal number |

Word documents a field result updated without `\* MERGEFORMAT` as taking its formatting from the field
code rather than from the previous result; LibreOffice behaves that way for both field forms. Before
#1115, a simple field was flattened to its cached text and kept the stale value on every page. The
converter now marks it like a complex field (`data-field`/`data-field-format`). The paginator replaces
the whole marked result (`substitutePageNumberFields`), so the cached run's formatting is
intentionally not reproduced. A `\* MERGEFORMAT` result that should keep prior formatting, and Word's
own output, still need a Word capture before Docxodus models them. Evidence: docx-fidelity-bench
`paragraphs-page-number-*` cases (bench issue #26).

## Fields: a deleted `PAGE` field keeps its code in `w:delInstrText` (issue #1272)

**Status:** fixed; Docxodus matches LibreOffice; Word not yet captured (October 2026)

A field deleted under tracked changes moves its instruction from `w:instrText` to `w:delInstrText`,
just as its text moves from `w:t` to `w:delText`:

```xml
<w:del w:id="31" w:author="Review Example" w:date="2026-09-12T10:00:00Z">
  <w:r><w:delText xml:space="preserve">Folio </w:delText></w:r>
  <w:r><w:fldChar w:fldCharType="begin"/></w:r>
  <w:r><w:delInstrText xml:space="preserve"> PAGE </w:delInstrText></w:r>
  <w:r><w:fldChar w:fldCharType="separate"/></w:r>
  <w:r><w:delText>1</w:delText></w:r>
  <w:r><w:fldChar w:fldCharType="end"/></w:r>
</w:del>
```

| Two-page footer, markup view | Word | LibreOffice | Docxodus before #1272 | Docxodus now |
|---|---|---|---|---|
| Deleted `Folio {PAGE} / {NUMPAGES}` | Not yet captured | `Folio 1 / 2`, `Folio 2 / 2` | `Folio 1 / 2` on both pages | `Folio 1 / 2`, `Folio 2 / 2`, struck through |
| Same footer, not deleted (control) | Not yet captured | `Folio 1 / 2`, `Folio 2 / 2` | `Folio 1 / 2`, `Folio 2 / 2` | unchanged |

`FieldRetriever.InstrText` read only `w:instrText`, and `WmlToHtmlConverter`'s revision wrappers
(`w:del`, `w:ins`, `w:moveFrom`, `w:moveTo`) converted their children run by run, so a field wholly
inside one was never grouped and never got its `data-field` marker. The field reader now reads
`w:delInstrText` for a deleted field, one whose begin character is inside a `w:del` or `w:moveFrom`
(a move source may also keep its code as `w:instrText`, as Word writes it). A live field still means
only its live `w:instrText`: one whose code was edited under tracking carries the old code as
`w:delInstrText` beside the new code as `w:instrText`, and a live field whose whole code was deleted
has no code at all. A wrapper that wholly contains a field groups it like paragraph content; a wrapper
inside an open field, or one a field crosses, does not, because the enclosing content already grouped
that field. Word often writes a deleted field as one `w:del` per run, each with its own id and
sometimes its own date. When every part of the field that shows text is the same kind of revision
(deletion or insertion), the marker is placed inside the revision that holds the result, because the
paginator replaces the marker's text and would otherwise drop the deletion styling. With deleted
content hidden, a deleted field shows nothing and gets no marker. The accepted view never sees the deleted
field: revisions are accepted before conversion. Evidence: docx-fidelity-bench `paragraphs-page-number-deleted-field` and its `-control`.

## Footnotes: how much of a page a long note may take (issue #1119)

**Status:** Docxodus matches LibreOffice; Word not yet captured (October 2026)

A note longer than a page continues on following pages. How much of each page it may use decides
where every later note line lands. Measured with LibreOffice 25.8.7.3 on one citing paragraph whose
footnote is a long run of exact-spaced lines (docx-fidelity-bench `notes-explicit-breaks-*`):

| Body height | Line height | Lines on a note-only page | Lines on the citing page |
|---|---|---|---|
| 468 pt | 4 pt | 107 | 107 |
| 648 pt | 4 pt | 150 | 150 |
| 864 pt | 4 pt | 202 | 202 |
| 1008 pt | 4 pt | 236 | 236 |
| 648 pt | 18 pt | 33 | 33 |
| 864 pt | 18 pt | 44 | 44 |

Every row fits one rule: the note container, including about 12.8 pt of separator and spacing,
may take at most 95% of the body area. The body keeps a twentieth even on a page with no body text.

| | Word | LibreOffice 25.8.7.3 | Docxodus before | Docxodus after |
|---|---|---|---|---|
| Citing page of the last body block | Not yet captured | ≤ 95% of body, less the body used | ≤ 60% | ≤ 95%, less the body used |
| Note-only continuation page | Not yet captured | ≤ 95% of body | ≤ 60% (five pages for a 90-line note) | ≤ 95% (three pages) |
| Body text still to place | Not yet captured | ≤ 95%: about 94% on page 5 of `TestFiles/NVCA-Model-COI.docx`, whose body continues on page 6 (measured for issue #1151) | ≤ 60% | ≤ 60% (unchanged) |

Docxodus applies the 95% limit through `noteBandLimit` and `MAX_EXHAUSTED_BODY_FOOTNOTE_AREA_RATIO`
in `npm/src/pagination.ts`, only once every body block is placed. The 60% cap that protects body
text still waiting for the page is unchanged, although LibreOffice lets notes take up to 95% there
too. Raising it to 95% leaves the NVCA charter's page count unchanged once its notes begin beside
their citations (issue #1151), and it changes the established layout of several paginator cases,
so it is left for its own change. Word's own limit needs a Word capture before either renderer's
value is treated as Word behaviour.

## Footnotes: a note begins on the page that cites it (issue #1151)

**Status:** Docxodus is within one page of LibreOffice on the NVCA model charter; Word not yet captured (October 2026)

A footnote whose citation sits near the bottom of a page may not fit in what is left of it. A renderer
can move the citing line to the next page, so that the note begins beside it, or keep the line and
start the note on the next page instead. Reproducer: `TestFiles/NVCA-Model-COI.docx`, 94 notes,
many of them longer than a third of a page, on 11pt Times New Roman body text.

| | LibreOffice 25.8.7.3 | Docxodus before #1151 | Docxodus after |
|---|---|---|---|
| Notes that begin on a later page than their citation | 1 of 94 | 23 | 0 |
| Pages holding notes and no body text | 0 | 2 (pages 6, 7) | 0 |
| Paginated view, total pages | 48 | 51 | 49 |
| Paginated editor (never splits a paragraph), total pages | — | 53 | 54 |

Measured with #1130's 6pt spacing between notes in place. Without it these changes gave 48 pages,
LibreOffice's count. LibreOffice lays out the same 6pt gaps, so they do not account for the
remaining page by themselves. The pages first part around pages 18–21: Docxodus carries a few lines
of a note onto the next page where LibreOffice has finished it, and from there trails it by up to a
page. That drift is not yet located.

LibreOffice's one exception is note 4. The first line of the "Fourth" paragraph cites it at the foot
of page 4, the notes already there fill the page, and LibreOffice keeps the line on page 4 and begins
the note on page 5. Docxodus moves that line to page 5, as it does for every other citation. Keeping
a note with its citation is the aim here; it has not been checked against a Word reference.

**What Docxodus did.** The paginator placed a paragraph whose lines fit and then, when one of its
notes could not start in the space left, deferred the whole note to the next page while the citation
stayed behind. The next page received the deferred notes before its own, ran out of room in turn,
and deferred more, until a page held nothing but notes. Two limits made this common:

- A paragraph that cites a note could not be split across pages at all, nor could a paragraph
  written as `h1`–`h6` (any Word paragraph whose style has an outline level, which in this charter
  is most numbered clauses), nor one that opens with a list number. Such a paragraph moved whole or
  not at all.
- A note paragraph that opens with a tab after its number (the FootnoteText style's
  `<w:tab w:val="left" w:pos="360"/>`, which 94 of the charter's notes use) could not be split
  either, so such a note either fitted whole or did not start.

**Fix.** `npm/src/pagination.ts`:

- `notesCanStartOnPage` answers whether a set of new notes can begin in the note band left beside
  a given amount of body: each fits whole except the last, which may split if its first line fits.
  When a block's notes cannot all start, the paginator splits the paragraph above the line whose
  citation fails (`tryFragmentParagraph`: a note-aware split breaks only where a visual line
  starts, so the citing line moves whole, with every citation on it), or, when the paragraph
  cannot be split because it keeps its lines together, moves it whole to the next page. The note
  is still deferred, as before, for a block that already starts its page, for a table (moving it
  whole would leave much of the page empty; splitting it at the citing row is not attempted), for
  a paragraph that a keep-with-next block before it keeps with (moving it would strand that block
  at the foot of the page), and in the paginated editor, which splits no paragraph: moving each
  citing paragraph whole there would leave more of every page empty than Word's split does. The
  editor's count rises by one page, from 53 to 54.
- `canFragmentParagraph` admits paragraphs written as `h1`–`h6` and paragraphs that cite notes; a
  head's notes are placed with the head.
- `leadingAtomicBoxes` lets a paragraph that opens with a list number, a note number or an empty
  tab box be split; no split point falls inside those boxes, so they stay whole with the first
  fragment.
- A paragraph is split for being too tall only when its own lines do not fit. When it is its notes
  that do not fit, the notes are split instead (earlier, every split left a paragraph that still
  could not take its notes whole, and it shrank word by word).
- Moving a citing line is worthwhile when a fresh page offers the note more room than is free on
  this one, which is the case when an earlier note's continuation already fills this page's note
  area. The test compares that free space, not the size of the area.
- The note cap judges only the notes a block adds. A block that cites nothing was refused once a
  carried continuation filled the capped area, so the page kept its notes and lost its body text
  (this also happened before #1151).

Tests: `npm/tests/pagination-note-citation-placement.spec.ts`.

## Paragraph spacing: consecutive numbered paragraphs keep their spacing after (issue #1151)

**Status:** Docxodus matches LibreOffice (October 2026)

```xml
<w:style w:type="paragraph" w:styleId="Heading2">
  <w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="4"/></w:numPr><w:spacing w:after="240"/></w:pPr>
</w:style>
<!-- three consecutive paragraphs in this style: FIRST, SECOND, THIRD -->
```

| Gap between FIRST and SECOND, NVCA model charter page 4 | |
|---|---|
| LibreOffice 25.8.7.3 | 24.7pt (one 12.65pt line and 12pt of spacing after) |
| Docxodus before #1151 | 12.5pt (no spacing) |
| Docxodus after | 24.64pt |

The markup's only property for removing the spacing between two paragraphs is
`w:contextualSpacing`, which applies between paragraphs of the same style (Word's List Paragraph
style sets it, which is why ordinary bulleted lists are tight). Nothing in these paragraphs sets
it, and LibreOffice keeps the spacing; Word has not been checked against a reference for this
document. The converter, following the original OpenXmlPowerTools behaviour, also
dropped the spacing after of every paragraph but the last in a run of paragraphs from one numbered
list. A numbered legal document lost its spacing between clauses: in the NVCA charter, 1. → 2.,
FIRST → SECOND → THIRD, "General" → "Voting", and the space after the "A. COMMON STOCK" heading.

**Fix.** `ApplyContextualSpacing` (formerly `GroupAndVerticallySpaceNumberedParagraphs`) in
`Docxodus/WmlToHtmlConverter.cs` groups paragraphs only by `w:contextualSpacing` and style. Numbering
no longer suppresses spacing. Restoring the spacing added two pages to the charter's paginated view
(measured before #1130),
height LibreOffice also lays out.

**Remaining gap.** LibreOffice also tightens list items whose spacing is `w:beforeAutospacing` /
`w:afterAutospacing` (HTML-style automatic spacing). Docxodus does not model automatic spacing for
list items; this predates #1151.

Tests: `Docxodus.Tests/NumberedParagraphSpacingTests.cs`.

## Footnotes: the note number is a run of the first paragraph, and spacing after stops at the margin (issue #1130)

**Status:** Docxodus matches LibreOffice; Word not yet captured (October 2026)

In OOXML a note's number is the `w:footnoteRef` run at the start of its first paragraph, so that
paragraph's indentation and spacing govern the number's line too:

```xml
<w:footnote w:id="1">
  <w:p>
    <w:pPr><w:ind w:left="360" w:hanging="360"/><w:spacing w:after="120"/></w:pPr>
    <w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>
    <w:r><w:t xml:space="preserve"> A note long enough to wrap onto a second line…</w:t></w:r>
  </w:p>
</w:footnote>
```

Measured with LibreOffice 25.8.7.3 on `TestFiles/NVCA-Model-COI.docx` (FootnoteText: 10 pt Times New
Roman, 6 pt after) and on two single-paragraph notes with `w:spacing w:after="120"`:

| | Word | LibreOffice 25.8.7.3 | Docxodus before | Docxodus after |
|---|---|---|---|---|
| Hanging-indented note, wrapped lines | Not yet captured | Start at the indent | Start at the band edge | Start at the indent (18 pt) |
| Number of a note split across pages | Not yet captured | Once, on the citing page | Once | Once |
| Gap between consecutive single-paragraph notes | Not yet captured | 6 pt (top-to-top 17.5 pt at an 11.55 pt pitch) | 0 | 6 pt |
| Spacing after the note area's last paragraph | Not yet captured | Not reserved: the last line ends on the bottom margin (720 pt) | Reserved only when the last note had several paragraphs | Not reserved (last line ends at 719.3 pt) |

Docxodus used to render the number beside the note text and force the first paragraph
`display: inline` to share its line; an inline paragraph ignores its own margins and text-indent.
`AddNoteNumberToFirstParagraph` in `WmlToHtmlConverter.cs` now puts the number inside the first
paragraph, and the generated rule
`.page-footnotes .footnote-item:last-child .footnote-content p:last-of-type` drops only the last
paragraph's spacing after. Word's treatment of that trailing spacing needs a Word capture.

---

## Paragraph borders on a right-to-left paragraph stay on their physical sides (issue #1154)

`w:ind`'s `w:left`/`w:right` on a `w:bidi` paragraph mean start/end: the leading indent moves to the
right. `w:pBdr`'s `w:left`/`w:right` do not move. ISO/IEC 29500-1 says the left paragraph border
"shall be displayed on the left side of the page around the specified paragraph. This shall not
change based on the paragraph direction" (quoted in the
[SDK's `LeftBorder` remarks](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.leftborder)).
The `RightBorder` remarks say the same of `w:right`. The only documented swap is the
`w:swapBordersFacingPages` compatibility setting, which swaps them on odd-numbered pages when the
document also uses `w:mirrorMargins` and `w:evenAndOddHeaders`.

```xml
<w:p><w:pPr><w:bidi/>
  <w:pBdr>
    <w:left w:val="double" w:sz="24" w:space="4" w:color="C0392B"/>
    <w:right w:val="single" w:sz="4" w:space="4" w:color="1F4E79"/>
  </w:pBdr>
  <w:ind w:start="1440" w:end="360"/>
</w:pPr><w:r><w:t>Right to left paragraph.</w:t></w:r></w:p>
```

| | Thick double (`w:left`) | Thin single (`w:right`) | Indents |
|---|---|---|---|
| ISO/IEC 29500-1 | left | right | start 72pt on the right, end 18pt on the left |
| LibreOffice 25.8.7.3 | left | right | mirrored, as the spec says |
| Docxodus | left | right | mirrored (issue #1129) |
| Word | not yet measured | not yet measured | |

Docxodus already behaved this way; `RightToLeftParagraphBorderTests` pins it. `CreateBorderDivs` in
`WmlToHtmlConverter.cs` writes each border on its own side and puts the leading and trailing indents on
the sides the paragraph's direction gives them, so each margin subtracts the outset of the border that
is physically on that side. A Word reference for this fixture would settle the last row.

### Mixed directions in one bordered group (issue #1155)

Consecutive paragraphs with identical borders and indents join into one box with `w:between` borders.
With unequal start and end indents, a left-to-right and a right-to-left paragraph have mirror-image
boxes, so they cannot share one.

| Two paragraphs, `w:ind w:start="720"`, second `w:bidi` | LibreOffice 25.8.7.3 | Docxodus before | Docxodus after |
|---|---|---|---|
| `w:end="720"` (equal) | one box, between border | one box | one box |
| `w:end="1440"` (unequal) | a box per paragraph, each mirrored | one box: one paragraph's indents on the wrong sides | a box per paragraph, each mirrored |

`BorderGroupKey` adds the direction to the key only when the box is not mirror-symmetric
(`BorderBoxIsMirrorSymmetric`), so grouping is otherwise unchanged. Word's grouping for this case has
not been measured.

## Exact-height rows: what the row hides (issues #1168, #1171)

An exact row (`w:trHeight w:hRule="exact"`) is as tall as its height, whatever its cells hold. A cell
can hide part of a line, a whole paragraph, and the border of a bordered group it cuts:

```xml
<w:tr><w:trPr><w:trHeight w:val="300" w:hRule="exact"/></w:trPr>
  <w:tc><!-- three paragraphs sharing one w:pBdr top+bottom group; 15pt holds one 11pt line -->
    <w:p><w:pPr><w:pBdr><w:top w:val="single" w:sz="8" w:space="1" w:color="FF0000"/>
      <w:bottom w:val="single" w:sz="8" w:space="1" w:color="FF0000"/></w:pBdr></w:pPr>
      <w:r><w:t>Alder clipped cell line.</w:t></w:r></w:p>
    <!-- Birch …, Cedar … with the same w:pPr -->
  </w:tc>
</w:tr>
```

| | Alder (fits) | Birch (cut) | Cedar (hidden whole) | Group's bottom border (below the row) |
|---|---|---|---|---|
| LibreOffice 25.8.7.3 | drawn | clipped; whole in the text layer | not drawn; not in the text layer | not drawn |
| Docxodus PDF export | drawn | clipped; whole in the text layer (#1140) | not drawn; in the text layer, after the visible text | not drawn (#1171; drawn below the row before) |
| Word | not yet measured | | | |

A header row (`w:tblHeader`) repeats at the top of each page the table continues on whatever its
`w:hRule` (reported for Word in issue #1171; not yet measured here). Docxodus repeated no header row,
whatever its rule, before that issue.

## Tables: nothing separates a table from the block above it (issue #1244)

Word starts a table's first row where the block above it ends. A paragraph's own after-spacing is the
only gap; with none (absent, zero or inherited as zero) there is no gap at all, and a table that
follows another table directly starts where that one ends. Compatibility mode 14 behaves the same.

```xml
<w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:t>WillowBefore boundary.</w:t></w:r></w:p>
<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>
  <w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>
  <w:tr><w:trPr><w:trHeight w:val="360" w:hRule="atLeast"/></w:trPr>
    <w:tc><w:p><w:r><w:t>WillowCell content.</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
```

Distance from the paragraph's line to the first cell's line, in points, with exact 16pt lines and
18pt minimum rows. Word is a Word for the web PDF export (October 23, 2026; no build reported);
Docxodus is the HTML converter's output laid out by Chromium:

| Block above the table | Word | Docxodus before | Docxodus after |
|---|---:|---:|---:|
| paragraph, after-spacing absent | 16.0 | 23.5 | 16.0 |
| paragraph, `w:after="0"` | 16.0 | 23.5 | 16.0 |
| paragraph, `w:after="0"`, table with a 1pt top border | 17.0 | 24.25 | 16.75 |
| paragraph, `w:after="120"` | 22.0 | 22.0 | 22.0 |
| paragraph, `w:after="240"` | 28.0 | 28.0 | 28.0 |
| paragraph, after-spacing 200 inherited from `w:docDefaults` | 26.0 | 26.0 | 26.0 |
| paragraph, `w:after="0"`, compatibility mode 14 | 16.0 | 23.5 | 16.0 |
| another table (first cell to second table's first cell) | 18.0 | 25.5 | 18.0 |

The converter used to give a table a 7.5pt top margin of its own whenever the block before it was a
paragraph without after-spacing, or anything other than a paragraph, to "separate" them (#108). Word
has no such rule, so the margin is gone. A floating table's `w:tblpPr/@w:topFromText` still becomes its
top margin.

Three differences remain, none part of this fix:

- When the second of two adjacent tables has a top border, Word does not add the border's width between
  them: its first line is still 18.0pt below the first table's. The HTML converter's output adds 0.75pt.
- The PDF export previously merged already adjacent tables during its `final` review-profile
  projection, dropping the second table's own properties (#1341). Revision acceptance now records
  source boundaries and joins compatible tables only when resolution creates their adjacency. Final
  and original exports therefore retain each table's own properties. The Word measurements above
  remain the original plain-HTML comparison; the fix adds native and production-export regressions.
  An original #1341 benchmark fixture adds a reference-engine caveat: LibreOffice 25.8.7.3 loads
  adjacent 4,800/6,000-twip tables with left cell margins 0/720 as one two-row table using the second
  table's width and margin. A paragraph-separated control stays as two tables. The saved DOCX and
  rendered PDF both confirm the difference. The new LibreOffice reference therefore is not proof of
  Word's table-boundary semantics; benchmark eligibility and full pixel measurements are retained.
  Native and export regressions assert each source table's properties independently. No new Word
  capture is claimed for this fixture.
- Word gives the exact zero-height rows of `tables-row-height-exact-zero-borderless` and
  `tables-row-height-exact-zero-only` (docx-fidelity-bench) the height of their content. In the second,
  the paragraph after the table starts 58.5pt below the one before it: four 14.6pt lines, one for that
  paragraph and one for each row. The #1339 fix treats an exact zero as unconstrained, matching the
  native Word captures; positive exact heights retain #1168's clipping. See the detailed entry below.

- Code: `Docxodus/WmlToHtmlConverter.cs` (`ProcessTable`).
- Tests: `HtmlConverterBlockSiblingTests`, `HC056_TablePrecededByParagraph_HasNoImplicitTopMargin`,
  `npm/tests/export-table-boundary.spec.ts`; bench cases `tables-boundary-*` (core-v51).

## Exact table rows: zero uses the content height (issue #1339)

`<w:trHeight w:val="0" w:hRule="exact"/>` is schema-valid, but Word for the web renders it at its
content height. It does not hide the row. Native PDFs captured on 2026-10-24 show this in borderless
and bordered rows, including a three-line cell. An exact row with its value omitted is likewise
unconstrained. Positive exact rows still clip: the separate Times New Roman control includes
one-twip empty rows and an 8pt row whose content exceeds its height. Those controls are not evidence
that an ordinary exact-zero row should disappear.

The [MS-OI29500 trHeight note](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/5919e0bd-e6ce-477e-8d66-0e5282f5c506)
(Part 1 §17.4.80, note c) distinguishes the standard, which permits exact zero, from Word's authoring
restriction against that combination. It does not prescribe Word's reader fallback. The content-height
fallback here follows the observed native Word PDFs. Note d also describes Word's bottom-padding
addition to a positive exact height; this fix does not change that separate geometry rule.

`ProcessTableRow` parses the height once and emits a height constraint and exact-row clip only for a
positive value. Zero therefore reaches pagination and export as ordinary cell content, with no clip
that would remove its PageMap fragments or PDF text. A positive height consumed entirely by borders
or padding can still produce a zero-area clip; synthetic PageMap tests retain that distinction.

`TestFiles/ZeroExactRows/word-observations.json` pins the two Times New Roman source packages,
native PDF hashes, fonts, pages and observed words. Source bytes are unchanged and validate against
Open XML SDK 3.5.1 Microsoft365. The earlier two Liberation Serif fixtures were also captured; Word
substituted Times New Roman, so those PDFs are diagnostic evidence rather than eligible font-matched
references. The source-identical Word downloads did not normalize the zero height. Word exposed no
exact build number. The PDF and UI evidence remain in the durable per-issue fixture pack.

Tests: `TableRowHeightRuleTests`, `TableRowHeightUnitTests`, `npm/tests/export-exact-row.spec.ts`
and `npm/tests/page-map-exact-row.spec.ts`.

## Footnotes: how much of a page the note area takes (issue #1178)

Four shapes, Letter page, 1in margins (648pt of body), Times New Roman 12pt single-spaced body
(13.75pt line pitch), 10pt single-spaced footnote text (11.5pt pitch), one line per note paragraph
unless stated:

| Shape | Word reference PDFs, page 1 / page 2 | LibreOffice 25.8.7.3, page 1 / page 2 | Docxodus at 60% (before) |
|---|---|---|---|
| a: line 1 cites a 90-line note, then 60 body lines | 1 body line and 53 note lines, from y 110.5 to 720.1 (the bottom margin) / 15 body lines, the note's last 37 lines | 2 body lines, 52 note lines / 14 body lines, 38 note lines | 19 body lines, 33 note lines |
| b: 20 body lines, line 21 cites the 90-line note | 21 body lines, 29 note lines / **no body**, 55 note lines from y 87.9 to 720.1; body resumes on p3 beside the note's last 6 lines | 21 / 2 body lines, 52 note lines | 21 / 19 body lines |
| c: lines 1-40 each cite a 4-line note | 11 body lines, notes to 720.1 / 10 body lines | 11 / 10 | 8 / 8 |
| d: line 1 cites one 70-line wrapping paragraph | 1 body line / 22 body lines | 2 / 21 | 19 / 19 |

Word holds no share of the band back for body text. A note takes everything below the line that
cites it; a carried note that does not end on a page takes the page alone (b, page 2), and the
body resumes on the page where it ends. A keep-with-next heading follows the same rule: when the
carried note ends on a page without room for the heading and its paragraph, the note keeps the page
and the chain opens the next one (Word reference PDFs of three such shapes; LibreOffice leaves the
heading alone at the foot instead). LibreOffice keeps some body beside a running note and holds a
twentieth of the band for the body once the body is placed (issue #1119).

Three more shapes, same setup:

| Shape | Word reference PDFs |
|---|---|
| line 1 cites a note holding an 80-row one-column table | p1: the citing line and rows 1-51; p2: rows 52-80 and the following body |
| line 30 cites a note holding a 30-row table | p1: 30 body lines and rows 1-18; p2: rows 19-30 |
| line 1 cites a note that is one 54-line keep-lines paragraph (too tall beside its citing line, short enough for a fresh band) | p1: the citing line and 30 body lines, no note; p2: the note alone |

So Word splits a table in a note between rows, keeps the citing line visible above it, and moves a
note it cannot start beside its citation whole to the next page, even onto a page that cites
nothing.

In `npm/src/pagination.ts`, `MAX_FOOTNOTE_AREA_RATIO` is 1 and `noteBandLimit` never gives the band
more than the body already placed on the page leaves; a page whose carried continuation does not
end on it is finished before any body block is placed on it (`continuationRunsOn`); and
`splitTableAtRows` splits a note's table at row-group boundaries in both the first part's and
the continuation's split.

A fitting nested table in a note exposed a separate conversion defect (#1217): the required
unformatted `<w:p/>` after the inner table had a one-twip line box but contained a full-size
synthetic nonbreaking-space run. Its descendant box protruded from the note area, so delivered
PDF verification rejected it as clipped. `WmlToHtmlConverter` now skips that placeholder for
paragraphs marked as structural table ends and gives the empty paragraph a one-twip minimum
height. Its source anchor remains measurable; explicit paragraph formatting and ordinary blank
paragraphs retain their normal rendering. Oversized rows and nested note tables use the separate
content-splitting path below.

For oversized note tables (#1217), `splitNoteTable` first tries complete row groups, then fills
an ordinary row's cells in one measured note band. It keeps complete paragraphs and keep chains
where possible, and uses the existing range fragmenter or recurses into an inner table when a
child cannot fit. Both the citation page and carried continuations use this path. Repeated
headers and column widths survive, and shared fragment identity reconciliation keeps one editor
anchor while retaining canonical source geometry on each page.

The converter exposes `w:cantSplit` and declared row-height rules as row metadata, including when
CSS classes are generated. Rows with those constraints, vertical merges or non-top alignment
retain the indivisible fallback; a header alone is not split as ordinary data. These cases need
separate layout rules. No clipping check or export verification rule is relaxed.

With paragraph fragmentation disabled in the editor, an oversized kept note paragraph is deferred
to a note-only page (#1217). The continuation path places that indivisible block once and advances,
preserving later body text and note order. The read-only split/clip fallback stays separate; this
editor placement rule does not claim that Word clips an oversized kept paragraph.

With paragraph fragmentation enabled, oversized kept note paragraphs split without clipping once
their keep can no longer be satisfied. Their first placement depends on compatibility mode. Matched
Word for the web exports captured on 2026-10-25 change only the settings part and its package
declarations; the body, note, styles and page geometry stay unchanged:

| Source | Compatibility 15 | Compatibility 14 or omitted |
|---|---|---|
| Two body lines; the first cites one kept 21-line note; 400 x 320pt page, 10pt margins, 15pt exact lines | Both body lines on p1; note alone on p2/p3, split 19/2 lines | Citation and 18 note lines on p1; following body and last three note lines on p2 |
| Thirty body paragraphs; the first cites a kept 1,200-word note; Letter page, 1in margins | All body on p1; note alone on p2-p5 | Citation and note start on p1; note continues alone on p2/p3; body resumes beside the note's end on p4, finishing on p5 |

The paginator therefore preserves the whole-note deferral in mode 15. Older modes relax an oversized
keep beside a fresh citation. Both paths retain the safe-inline fragmentation checks, source
identities, and widow control: the default or explicit on preserves two trailing lines; explicit off
permits one. A kept paragraph reaching an empty continuation band is split there rather than clipped
as a whole. The converter carries compatibility mode on staging as well as the body, because
embedding complete HTML in a viewer discards the body tag.

The natural-height probe also needs enough precision: Times New Roman measured at 1000px
produces a ratio of 1.149, versus its 2355/2048 metric ratio of 1.149902. At 10000px the probe gives
1.1499. The earlier underestimate accumulates through a long note and can admit one extra line.
Four native 120-line controls, with the citation first or mid-page and modes 14/15, assert every
body/note label's page. Mode 14 starts 53 lines beside a fresh citation, then carries 55 per full
note band. Mode 15 defers the note, then carries 55 per full band. Fresh note-only fragments use
the same measured physical band as carried continuations; the fixed whole-note guard is not also
subtracted from an already measured fragment's space. Text visibility and export clipping checks
remain required. This fixes the separately observed read-only kept-note clipping failure; it does not establish the missing original carried-note
"no measurable fragment" reproducer from #1217.

For a keep-with-next heading followed by a kept citing paragraph (#1217), the read-only path
tries the chain on a fresh page once. If its note still cannot start beside it, the heading may
remain while the paragraph moves; on a fresh page, that paragraph may relax keep-lines and split
before the citation. A relaxed paragraph keeps at least two visible lines on each side unless
`w:widowControl` explicitly disables the rule. Source keep metadata and canonical identities are
retained. The editor's paragraph-fragmentation switch remains authoritative.

Dense chains also distinguish a carried tail from a note that has finished beside a citation on
the current page. After such a note, a later kept citing chain can remain while its whole note
is deferred. A note already continuing from this page closes that option. Following body text
waits until the page where a newly split note ends, preserving body and note order.

A stock separator is a rule on an otherwise empty line. Its font is assembled with the document
styles and its shell reserves one natural `1lh`, resolved by the existing unrounded line-height
pass. The rule retains its gap above note text, with the extra reserve above the rule rather than
an added blank paragraph below it. Missing reserved definitions receive the same default layout;
custom separator stories with authored content retain their converted content. The separator is
presentation and never adds a note number, editor anchor or PageMap source fragment.

## Footnotes: a continued note stays last, and notes keep their order (issue #1179)

Five shapes, Letter page, 12pt body, single-line note paragraphs except in D and E, whose note 1 is
a single 20-line keep-lines paragraph (Word reference PDFs, LibreOffice 25.8.7.3, Docxodus's
paginated view since issue #1178):

| Case | Shape | Word | LibreOffice | Docxodus |
|---|---|---|---|---|
| A | `cites A` (90-line note 1), `cites B` (1-line note 2), 30 body lines | note 1 splits on p1; `cites B` moves to p2, note 2 follows note 1's end there | same | as Word |
| B | `cites A` (150-line note), page break, `cites B`, `cites C` (1-line notes) | p2 holds only note 1's continuation, no body; `cites B` and `cites C` open p3, where note 1 ends and notes 2, 3 follow in order | `cites B` alone on p2 beside the continuation; notes 2, 3 on p4 after note 1 ends | as Word |
| C | `cites A` (220-line note), then eight citing lines with plain lines between | p2-p4 hold only the continuation; body resumes on p5, where note 1 ends, and notes 2-9 follow in order | `cites X2` on p2, only `plain2 2` on p3, `cites X3` on p4; note 1 ends on p5 and notes 2-9 start there | as Word |
| D | 40 filler lines, a paragraph citing note 1, `cites B` | the citing line moves to p2 with its note | same | same |
| E | 40 filler lines, a one-row table citing note 1, `cites B` | the table moves to p2 with its note and `cites B` | the table stays on p1; note 1 and note 2 go to p2, in order | as Word (#1217) |

All three keep a continued note last on its page and every note in citation order. While the
paginator still capped the note area at 60% (until issue #1178), Docxodus laid B and C out like
LibreOffice, with a citing line beside each continuation, and A with `cites B` beside the
continuation on p2 and note 2 after note 1's end on p3. Where Word and
LibreOffice differ, Word gives a page whose carried note runs on to the next page entirely to that
note (B, C; see the note-area entry for issue #1178), and moves a table whose note cannot start
beside it to the note's page (E).

For E, the citation-placement path also handles tables in either paragraph-fragmentation mode
(#1217). It reuses `splitTableAtRows` to leave complete earlier row groups on the current page,
without cutting vertical merges or leaving repeated headers alone. When no prefix fits, the
whole table moves only if the fresh body band can hold it and start its notes. The keep-chain path
still owns a preceding heading that keeps with the table. Captions and nested tables remain
indivisible. The existing `notes-keep-lines-note-in-table` benchmark and a new fitting-note control
check the delivered PDF, complete text inventory and note PageMap; synthetic browser tests also
cover merged rows, repeated headers, captions and oversized-note progress.

The Docxodus rule is `laterNotesClosed` in `PaginationEngine.flowToPages`
(`npm/src/pagination.ts`): a page closes to new notes once one of its notes continues
(`nextPageContinuation`) or a note cited on it is deferred (`deferredFootnoteIds`), and
`prepareCurrentContinuation` moves the notes deferred to a page on with a carried note that runs on
again; `finishPage` hands them back when the last body page lets the carried note end.

## Splitting a paragraph: the new paragraph mark ends the first half (issues #1297, #1298)

Word's tracked Enter marks the pilcrow that now ends the text *before* the cursor as inserted.
`TestFiles/RP/RP006-Inserted-Paragraph-Mark.docx` shows the shape:

```xml
<w:p>
  <w:pPr><w:rPr><w:ins w:id="0" w:author="Eric White" w:date="2017-03-24T21:58:00Z"/></w:rPr></w:pPr>
  <w:r><w:t xml:space="preserve">Video provides a powerful way to help you prove your point. </w:t></w:r>
</w:p>
<w:p><w:r><w:t>When you click Online Video, …</w:t></w:r></w:p>
```

So the original mark goes on ending the second half, together with what belongs to the mark
alone: its own revision (`w:pPr/w:rPr/w:ins` or `w:del`) and, for a paragraph that ends a section,
the section's `w:sectPr` (a copy on the first half would add a section). Ordinary paragraph
properties apply to both halves.
Rejecting an inserted mark joins its paragraph into the next one. Which paragraph's properties
the joined paragraph keeps differs between implementations:

| Reject an inserted mark between P1 (`jc=center`, mark inserted) and P2 (`jc=right`) | Joined paragraph keeps |
|---|---|
| Word | unmeasured: RP006 and RP008 have the same properties on both sides, so they cannot tell |
| LibreOffice (measured in review with the repository's headless listener) | P1's `w:pPr` (`center`) |
| Docxodus (`RevisionOps.MergeParagraphWithNext`; `RevisionProcessor`'s coalescing takes the last member's `w:pPr`) | P2's `w:pPr` (`right`) |

The split design gives the same answer either way. The first half keeps the original paragraph's
properties, the ones LibreOffice keeps. Whatever the split changed on the second half, the one
Docxodus keeps, is recorded as the editor's own property change, which reject-all undoes. So on
every row of issues #1297 and #1298, reject-all gives back the original paragraph in both.

| Split of a paragraph whose mark Alice inserted | Mark revision on | Reject everything |
|---|---|---|
| Copied onto both marks (Docxodus before #1297) | both, one shared `w:id` | ambiguous; nothing resolves individually |
| Left on the first half's new mark | first half | undoes the user's untracked split, keeps Alice's break |
| On the second half's original mark (Docxodus now) | second half | keeps the untracked split, removes Alice's break |

The same rule explains Enter at the end of a heading: the empty second paragraph holds the
heading's original mark in the heading's next style. Rejecting Alice's insertion of that mark
joins the empty paragraph into the following one, and the heading keeps its style. Under
tracking, Docxodus records whatever the split changed on the second half (the page break left
on the first half, the next style, the dropped mark formatting) as the editor's
`w:pPrChange`/`w:rPrChange`. Rejecting everything then restores the original paragraph's
properties. A pending property change the paragraph already had keeps its baseline instead, and
it survives the next-style rebuild for that reason; otherwise the editor's change would record the
heading as the baseline, and reject would keep the heading that the pending change should undo.

Code: `DocxSession.SplitParagraph`, `TrackSplitOffProperties`, `RemoveMarkOwnedProperties`.
Tests: `DocxSessionClonedRevisionTests` (RID003, RID004, RID007, RID010, RID011, RID020, RID022).

## Merging paragraphs under tracking: the deleted mark and whose properties survive (issue #1313)

A tracked join (Delete at a paragraph's end, Backspace at the next one's start) is the inverse of
the tracked Enter above: the mark that ends the first paragraph is marked deleted, and both
paragraphs stay until review.

```xml
<w:p>
  <w:pPr><w:jc w:val="center"/><w:rPr><w:del w:id="7" w:author="Bob" w:date="…"/><w:b/></w:rPr></w:pPr>
  <w:r><w:t>First.</w:t></w:r>
  <w:ins w:id="8" w:author="Bob" w:date="…"><w:r><w:t xml:space="preserve"> </w:t></w:r></w:ins>
</w:p>
<w:p>
  <w:pPr><w:jc w:val="center"/>
    <w:rPr><w:b/><w:rPrChange w:id="9" w:author="Bob" w:date="…"><w:rPr/></w:rPrChange></w:rPr>
    <w:pPrChange w:id="10" w:author="Bob" w:date="…"><w:pPr><w:jc w:val="right"/></w:pPr></w:pPrChange></w:pPr>
  <w:r><w:t>Second.</w:t></w:r>
</w:p>
```

Accepting the deleted mark joins the first paragraph into the second. The question is again which
paragraph's properties the joined paragraph keeps:

| Accept a deleted mark between P1 (`jc=center`) and P2 (`jc=right`), no property change recorded | Joined paragraph keeps |
|---|---|
| Word | unmeasured |
| LibreOffice | unmeasured |
| Docxodus (`RevisionOps.MergeParagraphWithNext`) | P2's `w:pPr` (`right`), the surviving mark's |
| Untracked `DocxSession.MergeParagraphs` | P1's formatting (`center`); P2's section boundary |

The tracked merge does not depend on the answer. It writes P1's properties onto P2 as the editor's
own `w:pPrChange`, and P1's mark formatting as the editor's `w:rPrChange` on P2's mark. The only
row the encoding needs is the one Docxodus measures: accepting everything gives P1's properties,
which is what an untracked merge gives. A consumer that keeps P1's properties on accept gives the
same result. Rejecting everything resolves the property changes and restores the mark, so both
paragraphs come back exactly. What belongs to P2's mark alone (its own revision, its section break)
is not touched, so a section break that P2 ends stays where it was.

A tracked merge whose first mark ends a section remains unsupported. An untracked merge removes
that boundary and keeps the second mark's entire `w:sectPr`, including its pending section-property
change, if present (#1350). Previously it kept the first boundary and discarded the second; merging
two paragraphs before the next section could silently make both sections use the final section's
layout. The first paragraph's other properties still survive. Word and LibreOffice section-boundary
behaviour has not been measured for this reproducer.

The copy of P1's properties is one value, so it cannot follow a pending change to them. Say Alice
centred P1 (`w:pPrChange`, before: `left`) and Bob then merges. Bob's copy is `center`. Reject
Alice's change and accept Bob's merge, and the joined paragraph comes out `center`, the formatting
Alice's rejection was meant to remove; rejecting her change first and then merging untracked gives
`left`. So a tracked merge is refused (`UnresolvedStructuralRevision`) while P1 or its mark has a
pending property change, whoever made it. A pending change on P2 is refused too: the merge would
fold its own change into it, so one revision would record two edits, and a change of Alice's the
merge took back to its baseline would simply be gone. The one exception is the editor's own change
that the merge takes back to its baseline, the Enter-then-Backspace case, which is resolved.

Code: `DocxSession.MergeParagraphs`, `RefuseTrackedMerge`, `TrackMergedProperties`.
Tests: `DocxSessionTrackedEditGapTests` (`TrackedMerge_*`, `TrackedSplitThenMerge_*`,
`TrackedEnterAfterAHeadingThenBackspace_*`).

## Splitting a paragraph inside a field's result (issue #1313)

```xml
<w:p><w:r><w:t xml:space="preserve">See </w:t></w:r>
  <w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> REF x </w:instrText></w:r>
  <w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>Clause twelve</w:t></w:r>
  <w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t xml:space="preserve"> now</w:t></w:r></w:p>
```

Word lets a field's result span paragraphs: a table of contents is one field across many. But a
split at an offset inside "Clause twelve" splits the field's displayed result, not the field.
Before #1313, Docxodus put `begin`, the instruction, `separate` and "Clause" in the first paragraph,
and " twelve" with the `end` mark in the second. The same thing happened at offset 17, right after
the result. The `end` mark has no text, so it sits at offset 17 too, and every zero-width child at
the split offset moved to the second paragraph.

| Split at | Docxodus before | Docxodus now |
|---|---|---|
| 4 (before the field) | field whole in the second paragraph | same |
| 10 (inside the result) | field cut in two | refused, `UnsupportedInlineBoundary` |
| 17 (after the result) | field cut: `end` alone in the second paragraph | field whole in the first paragraph |

Only an offset strictly between a field's begin and end marks is refused, and only for a field
whose begin and end are both in the paragraph. A field that already runs across paragraphs is not
the paragraph's to keep whole, so a split inside it goes ahead. Word may write a mark in the same
run as text on the far side of an edge (`<w:r><w:fldChar w:fldCharType="end"/><w:t> now</w:t></w:r>`),
or inside the same hyperlink or tracked insertion as that text. A split at that edge first cuts the
run and each such container at the mark, as a text split cuts them, so the field stays whole on its
side. The check splits a copy of the paragraph and compares where each field's marks land, so a
field inside a tracked insertion or a hyperlink is covered as well; an edge inside a container a
split keeps whole (a content control) is refused like the inside.

Code: `DocxSession.SplitWouldCutAField`, `ParagraphFields`, `SeparateFieldMark`, `MoveInlineChildrenAfter`.
Tests: `DocxSessionTrackedEditGapTests` (`Split*Field*`).

## Hidden paragraph mark: which paragraph's keep properties govern the joined block (issue #1294)

A paragraph whose mark is hidden (`w:pPr/w:rPr/w:vanish`) is joined to the paragraph after it, the
usual run-in heading construction. Minimal reproducer (Letter page; headings repeated so the
sections sit at the end of a page):

```xml
<w:p><w:pPr><w:keepNext/><w:keepLines/><w:rPr><w:b/><w:vanish/></w:rPr></w:pPr>
  <w:r><w:t>Term.</w:t></w:r></w:p>
<w:p><w:pPr><w:keepLines/></w:pPr><w:r><w:t xml:space="preserve"> The clause text.</w:t></w:r></w:p>
```

Three such sections, with room for two on page 1 (benchmark case `paragraph-mark-hidden-run-in-keep`):

| | Section 1 | Section 2 | Section 3 |
|---|---|---|---|
| Behaviour Docxodus #1294 specifies | page 1 | page 1 | page 2 |
| LibreOffice 25.8.7.3 | page 2 | page 2 | page 2 |
| Word | pending: Word reference PDFs available here cannot adjudicate a hidden-mark join | | |
| Docxodus before the fix | chained onto page 2 as one group (the block took the heading's `keepNext`) | | |

**Rule.** The joined block takes `data-keep-with-next` and `data-keep-lines` from its last paragraph.
A heading's `keepNext` is satisfied by the text it runs into, so the closing paragraph's mark is the
one that governs what follows the block, including an explicit `keepNext`/`keepLines` off there.
`data-page-break-before`, which belongs to the block's start, still comes from the first paragraph.
This is the behaviour the issue specifies; it has not been checked against Word.

**Code.** `WmlToHtmlConverter.AddPaginationDataAttributes`, called from `ProcessParagraph` with the
first and last paragraph of the join. Tests: `HiddenParagraphMarkTests.RunInBlock_*`.

## Revisions: how Word prints All Markup with balloons (issue #1300)

**Reproducer.** Track one bold change and one indent change:

```xml
<w:p><w:r><w:rPr><w:b/><w:rPrChange w:id="1" w:author="Alice" w:date="2025-01-02T09:00:00Z"><w:rPr/></w:rPrChange></w:rPr><w:t>buyer</w:t></w:r></w:p>
<w:p><w:pPr><w:ind w:left="720"/><w:pPrChange w:id="2" w:author="Alice" w:date="2025-01-02T09:00:00Z"><w:pPr/></w:pPrChange></w:pPr><w:r><w:t>Indented clause.</w:t></w:r></w:p>
```

Then print or save as PDF with All Markup and "Show only comments and formatting in balloons".

| | Word | Docxodus before | Docxodus now (`commentColumnPlacement: "balloons"`) |
|---|---|---|---|
| Paper | Source size. The page is shrunk beside a 3.7in grey area. | Margin column inside the page, or an extra panel beside it | As Word |
| Bold change | `Formatted: Font: Bold` balloon | Change bar; `title="Bold added"` | As Word |
| Indent change | `Formatted: Indent: Left: 0.5"` balloon | Change bar; `title="Paragraph formatting changed"` | As Word |
| Insertions and deletions | Inline | Inline | Inline |

**Measured from Word-printed All Markup PDFs.**

- **Shrinking.** Word shrinks the page as far as its text column's right edge plus the 266.4pt
  (3.7in) area, to the paper's width. The shrunken page stands at the paper's foot, 10.8pt above it.
  The grey area starts 9pt right of the text column.
- **Balloons.**
  - The left edge is 31.6pt past the text column.
  - Borders are 0.5pt in the author's colour, with 3pt corners.
  - Balloons are 1pt apart.
  - Labels and wording use the Balloon Text style, else Segoe UI 9pt; comment text keeps its own
    font.
- **Leaders and marks.** Hairline leaders run along the anchor's line, then to the balloon.
  Formatting changes are marked with a small triangle.
- **Colours.** Authors take `#0078D4`, then `#D13438`. Comment authors are numbered on their own.
  The change bar sits 30pt left of the text.
- **One balloon per formatting revision.** Runs that share the wording and the recorded properties
  get a single balloon. A move to another level of the same list is undescribed. A tracked paragraph
  mark between two such runs ends the revision (reported from a Word print of a document whose
  tracked highlighting ran across paragraphs, some with tracked marks; not re-measured here). The
  mark stands at its paragraph's end, so the marks between the runs are the earlier run's own
  paragraph mark and those of whole paragraphs between them. The later run's paragraph mark follows
  the run and does not end the revision; Docxodus 14.2.0 ended it there and started a new balloon.

**Docxodus.** `Internal/WordFormatWording.cs` writes the wording, `npm/src/markup-balloons.ts` lays
the balloons out, and `docs/architecture/all_markup_layout.md` gives the design.

### When a page's balloons cannot all fit beside it

**Reproducer.** One Letter page of short paragraphs, each carrying a long comment, 25 of them. In the
measured document each comment filled eight lines of its balloon (about 430 characters), and one
comment ran to several paragraphs with empty ones between. Each comment's `w:date` (the author's
clock) differs from the UTC date the modern comment parts record:

```xml
<!-- document.xml, repeated for ids 0 to 24 -->
<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>Section 1 applies to the transferred assets.</w:t></w:r>
  <w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>
<!-- comments.xml -->
<w:comment w:id="0" w:author="Ines Moreau" w:initials="IM" w:date="2025-06-11T16:05:00Z">
  <w:p w14:paraId="10000000"><w:r><w:t>Note 0: the warranty period … (a few hundred characters)</w:t></w:r></w:p>
</w:comment>
<!-- commentsIds.xml maps paraId 10000000 to durableId 20000000; commentsExtensible.xml: -->
<w16cex:commentExtensible w16cex:durableId="20000000" w16cex:dateUtc="2025-06-11T21:05:00Z"/>
```

`printOverflowFixture()` in `npm/tests/docx-markup-balloons-fixture.ts` builds this document with
shorter comments (about six balloon lines each). Print it with All Markup and comments in balloons,
once with Save as PDF and once to a PDF printer.

| | Word, Save as PDF | Word, Print to PDF | Docxodus 14.2.0 | Docxodus since PR #1392 and PR #1393 |
|---|---|---|---|---|
| Which balloons are cut | Balloons stay whole from the top while the rest, each cut to one line, still fit at or below their own lines: 3 whole, 22 cut | Same | Lines come off the tallest balloon until the stack fits, so the longest comments all end part way | Word's rule; on the measured document 4 whole, see Residuals |
| A cut balloon | First line only, 16.1pt tall at 9pt, ending in a small white box with a hairline border holding `…` | Same box, reading `… [n]` | Lines cut from the tallest, ending in a text ellipsis | As Word |
| The rest of a cut comment | Printed nowhere | Printed in full on pages after the last, under a bar naming page, number, label, author and date | Printed nowhere | `balloonOverflow: "truncate"` (default) prints nowhere; `"appendix"` prints the full text as Print to PDF does, without Word's leaders (see Residuals) |
| The date in that bar | — | `w:date`'s digits, not the UTC date | — | As Word |

LibreOffice: not compared; it has no equivalent balloon print layout.

**Why.** Word gives the space to the earliest balloons, and keeps every balloon at or below its own
line rather than pulling the stack up to keep one more whole. It treats the full text as a printing
matter: the file-saving route drops it, the printer route appends it.

**Docxodus outside the measured case.**
- A stack the page could hold if pulled up from its foot is pulled up and printed whole, unchanged
  from 14.2.0. This is assumed, not measured: in the measured case the stack could not fit even
  pulled up.
- If even the cut balloons cannot fit, they run off the page's foot and are cut there. Such a balloon
  has no box, but `"appendix"` still lists its comment.
- Every cut balloon is reported in the export's `comment_balloon_truncated` warning.

**Residuals.**
- On the measured document Docxodus keeps 4 balloons whole where Word keeps 3. Measured on a host
  without Segoe UI and Times New Roman, the substitutes make an eight-line balloon 87.9pt tall
  against Word's 88.5pt. Four whole balloons then end 2.9pt above the page's foot, where Word's
  would end within about a point of it.
- `"appendix"` omits three things Word's Print to PDF also prints: a leader from each comment out
  past the paper's edge, a mark in the left margin, and a final page holding only the last leader.
  So Word's PDF has one more page.
- Not measured yet:
  - whether Word pulls up a stack that fits only that way;
  - what Word does when even its cut balloons cannot fit;
  - whether the appendix bar names the physical or the displayed page number;
  - whether page-count fields count the appendix pages (Docxodus counts only the document's own).

**Code.** `fit()` and `cutShort()` in `npm/src/markup-balloons.ts`, and `npm/src/balloon-appendix.ts`.
**Tests.** `npm/tests/markup-balloons.spec.ts` and `npm/tests/export-markup-balloons.spec.ts`; the
tests named for cut balloons, the pull-up case and the appendix.

## Comments in a footer or footnote (issue #1310)

OOXML permits a comment anchor outside the body. Docxodus preserves it under its documented
export contract, even though Word ignores these non-body anchors. Minimal
reproducer, a footer holding the range and its reference (the footnote form is the same markup inside
`w:footnote`):

```xml
<w:ftr><w:p>
  <w:commentRangeStart w:id="5"/><w:r><w:t>OakFooterLine Ledger copy.</w:t></w:r>
  <w:commentRangeEnd w:id="5"/><w:r><w:commentReference w:id="5"/></w:r>
</w:p></w:ftr>
```

| | Footer or footnote comment with a reference | The same range with no `w:commentReference` | Body comment with no reference |
|---|---|---|---|
| Word | Microsoft documents that these non-body anchor elements are ignored; no new capture | same documented restriction | no new capture |
| LibreOffice 25.8.7.3 | will not open the package: `soffice --convert-to pdf` reports "source file could not be loaded" | opens; draws no comment | opens; draws no comment |
| Docxodus before the fix | no highlight, marker or entry, in every comment mode; no export warning | the same | highlight and note |
| Docxodus now | highlight, marker and entry, as a body comment | highlight and note | unchanged |

**Analysis.** Microsoft's [commentReference implementation note](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/84a35ef3-9c12-4b59-ab4c-4a88f520d06c)
and [commentRangeStart note](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/64b95366-980f-4c75-9519-b745345ded9a)
say that OOXML permits these anchors but Word ignores them outside the body. Preserving their
content here follows Docxodus's export contract; it does not claim Word parity.
LibreOffice's failure is a load failure, not a rendering choice. It reproduces with the
reference run on its own (no range) and with a `CommentReference` character style; removing only the
reference run makes the package load. A comment with no reference has no annotation in LibreOffice,
so it prints nothing even in the body. Docxodus lost the comment because its comment state was
attached to the body's part only; a range in another story found none.

**Code.** `WmlToHtmlConverter.ConvertToHtml` annotates a per-story `CommentTracker` onto every story
root and each individual note; `npm/src/pagination.ts` places a footer comment's margin note.
The nearest tracker scopes open ranges to one note, so an open range cannot highlight its peer.
A rendered range without a reference also keeps its entry when another story contains a real
reference; selecting only reference-mark IDs previously discarded that note.

**Validation.** Open XML SDK 3.5.1 reports `Sem_MissingReferenceElement` for permitted non-body
comment anchors even when the matching comment exists in the main document's comments part.
The SDK disallows adding that comments part directly to a header/footer. Tests retain only this
bounded semantic finding for these four story part types and verify every target id against the
actual comments part; all other validation errors fail. This does not suppress renderer or export
warnings. Benchmark metadata retains each SDK finding verbatim. Design:
`docs/architecture/comment_rendering.md`, "Comments outside the body". Tests:
`StoryCommentRenderingTests`, `HtmlConversionEditorRenderTests.HER007`,
`npm/tests/export-story-comments.spec.ts`. The paired fidelity-bench story cases use the range-only
form, because the LibreOffice reference track cannot open the other.

## Contributing

When adding new corner cases to this document:

1. **Provide a minimal reproducer**: Include the relevant XML snippets and a description of how to reproduce
2. **Document all renderers**: Test in Word, LibreOffice, and Docxodus
3. **Reference the spec**: Link to relevant ECMA-376 sections
4. **Identify the code**: Point to the specific Docxodus files/functions involved
5. **Propose a fix**: If possible, outline how the issue might be resolved

### Different default paragraph style IDs in a comparison

A style's display name does not establish its identity. Two packages can each name their default
paragraph style `Normal`, while giving it different IDs and different fonts or spacing. An implicit
paragraph uses its own package's default; copying it into a left-based comparison without an
explicit reference would select the left default in both revision views.

Minimal synthetic inputs have identical `<w:p><w:r><w:t>Jade beacon remains bright.</w:t></w:r></w:p>`
content, and these distinct default definitions:

```xml
<!-- Original -->
<w:style w:type="paragraph" w:default="1" w:styleId="BaseText">
  <w:name w:val="Normal"/><w:pPr><w:spacing w:after="0" w:line="240"/></w:pPr>
  <w:rPr><w:rFonts w:ascii="DejaVu Sans"/><w:sz w:val="20"/></w:rPr>
</w:style>
<!-- Revised uses FreshText, DejaVu Serif, size 28, after 360, line 360. -->
```

| View | Word | LibreOffice | Docxodus before #1050 | Docxodus after #1050 |
|---|---|---|---|---|
| Accepted font / size / after / line | Not measured | Not measured | Sans / 20 / 0 / 240 | Serif / 28 / 360 / 360 |
| Rejected font / size / after / line | Not measured | Not measured | Sans / 20 / 0 / 240 | Sans / 20 / 0 / 240 |

`IrMarkupRenderer` retains the left default definition, imports the revised definition without its
package-wide default flag, and adds `w:pStyle` to revised implicit consumers. A paired paragraph's
`w:pPrChange` archives its original reference (or its absence). Acceptance keeps the revised
reference; rejection restores the original selection. Insertions and move destinations disappear
on rejection and need only the revised reference. This uses default roles rather than matching
display names, so unrelated custom styles named `Normal` remain separate. Synthetic regression
tests cover both explicit and implicit references, inheritance, overrides, structural edits,
accepted/rejected text, native property revisions, and Office 2019 schema validation.

### A new heading borrowing a body paragraph's leading phrase

An anchored replacement region can contain both a new heading and a close edit of the original
body paragraph. Counting shared content words alone can tie those two candidates, even though
one is a much closer correspondence. A fully synthetic example is:

```text
Original body: Bronze radar channel cabinet enabled.
New heading:  Bronze radar channel cabinet service calendar for the coastal crew.
Revised body: Bronze radar channel cabinet disabled.
```

The heading has a paragraph style with `w:keepNext`, `w:outlineLvl w:val="1"`, and a 14pt run
size. Both candidates share the four leading content words with the original body. The heading
adds several unmatched concepts, while the revised body changes one word.

| Alignment | Word | LibreOffice | Docxodus before #1052 | Docxodus after #1052 |
|---|---|---|---|---|
| Original body counterpart | Not measured | Not measured | New heading | Revised body |
| New heading classification | Not measured | Not measured | Modified body | Inserted paragraph |
| Accepted/rejected paragraph text | Not measured | Not measured | Correct | Correct |

`IrBlockAligner.SameSlotPair` retains its positional preference when evidence is equal. A still-free,
non-crossing candidate can now outbid a same-slot pairing when it shares the same content-word
count but has strictly stronger content-word Jaccard overlap. With target and shared count fixed,
this means fewer unmatched distinct content words. The rare-word index probes one additional word
for a tie and considers only shorter candidate vocabularies. Genuine heading edits and paragraph
style changes can still pair; style names and font sizes are not a blanket matching prohibition.
Tests cover styled and unstyled insertions, genuine heading edits, body edits, paragraph boundaries,
splits/merges, unrelated rewrites, native inserted paragraph marks, both endpoints, schema validity,
and agreement between the indexed tie decision and exhaustive search.

### Document-default changes with table consumers

**Status:** Fixed (issue #1051, October 2026)

A literal `w:docDefaults` font/spacing change with otherwise identical style definitions used to
be lost whenever either input contained a table. The paragraph-only style projection deliberately
excluded tables: putting defaults into paragraph styles would raise their precedence above conditional
table formatting. Removing that gate alone would produce the wrong appearance.

Minimal synthetic input (same `Normal` definition and newly invented body/cell sentences on both sides):

```xml
<w:docDefaults>
  <w:rPrDefault><w:rPr><w:rFonts w:ascii="DejaVu Serif" w:hAnsi="DejaVu Serif"
    w:eastAsia="DejaVu Serif" w:cs="DejaVu Serif"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:rPrDefault>
  <w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="360" w:line="360" w:lineRule="auto"/></w:pPr></w:pPrDefault>
</w:docDefaults>
```

The original declares DejaVu Sans, sizes 20, after 0, and line 240. Add an unchanged one-cell table
and the old accepted result retains those original values in both body and cell paragraphs.

| Consumer | Word | LibreOffice | Docxodus before | Docxodus after |
|---|---|---|---|---|
| Body and cell default font/spacing | Not measured | Not measured | Original defaults survive acceptance | Acceptance uses revised defaults; rejection restores original |
| Conditional first-row after/font override | Not measured | Not measured | Unchanged, but other inherited axes remain old | Override remains authoritative; other axes follow revised defaults |

`IrDocDefaultsConsumerProjection` resolves detached right sources with both defaults sets through
`FormattingAssembler`, materializes only affected literal font/size/spacing axes, and carries the old
direct properties in native `pPrChange`/`rPrChange`. The package keeps left defaults and style definitions.
The projection includes paragraph marks, including empty paragraphs. Inserted marks omit formatting
history so rejecting it cannot erase their structural insertion marker. Source slices receive unique
native revision IDs after emission; cached IR snapshots retain their original sources.

Synthetic tests cover conditional and direct precedence, inherited styles, edited text, inserted/deleted
paragraphs, row insertion, splits/merges, removed declarations, snapshot reuse, and Office 2019 validation.
The existing paragraph-only control remains green. Other package presentation changes and the original
unsupported consumer gates keep their existing behavior.

### Shared table-style cell margins in comparisons

**Status:** Fixed (issue #1053, October 2026)

The comparison style updater tracks paragraph/run properties, but table-style geometry has no native
style-level property history. A shared table style's `w:tblCellMar` therefore remained at its original
values even after accepting all revisions. The equivalent direct-table change already worked.

Minimal synthetic style payload (all body/cell wording is newly invented):

```xml
<w:style w:type="table" w:styleId="BeaconGrid">
  <w:name w:val="Beacon Grid"/>
  <w:tblPr><w:tblCellMar>
    <w:top w:w="300" w:type="dxa"/><w:left w:w="120" w:type="dxa"/>
    <w:bottom w:w="180" w:type="dxa"/><w:right w:w="120" w:type="dxa"/>
  </w:tblCellMar></w:tblPr>
</w:style>
```

The original uses top/bottom 0. Both packages reference the same style on an unchanged one-cell table.

| Change | Word | LibreOffice | Docxodus before | Docxodus after |
|---|---|---|---|---|
| Shared style top/bottom 0/0 → 300/180 | Not measured | Not measured | Acceptance retains 0/0 | Acceptance 300/180; rejection 0/0 |
| Direct table-margin control | Not measured | Not measured | Reversible | Reversible |
| Conditional first-row top 75 → 95, direct bottom 65 | Not measured | Not measured | Conditional style change lost | Acceptance 95/65; rejection 75/65 |

`IrTableStyleMarginProjection` keeps left table-style definitions and resolves right consumers under
both the retained and revised style graphs. It expresses table defaults through native `tblPrChange`
and changed cell/conditional exceptions through native `tcPrChange`, preserving each original direct
shell for rejection. Imported children still need projection when their parent is a retained shared
style. Unchanged same-URI header stories receive projected consumers when their original and carried
XML match; edited stories keep the existing relationship-aware emitter.

`FormattingAssembler` now merges margin axes independently, as required when a side is omitted from
a higher-precedence layer. See [the SDK's ISO/IEC remarks for top margins](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.topmargin?view=openxml-3.0.1).
`WordprocessingMLUtil` orders the merged margin children as well as the table/cell shells, so a direct
bottom margin followed by a newly projected top margin remains valid OOXML.

Synthetic regressions cover direct table/cell overrides, inherited and imported child styles,
conditional regions and changes, unrelated tables, edited/nested/inserted tables, row insertion,
removed declarations, logical margin axes, unchanged headers, combined document-default changes,
accept/reject text and geometry, and Office 2019 validation. The implementation reuses the detached
consumer sources introduced by the document-default fix rather than changing cached IR facts.

### One style id used for styles of different types in a comparison (issue #1184)

A `w:styleId` names one style across every type: two `w:style` elements may not share an id even when
one is a paragraph style and the other a character style (the Open XML validator reports
`Sem_UniqueAttributeValue` on `/word/styles.xml`). Two documents can still each use the same id for
styles of different types:

```xml
<!-- original -->
<w:style w:type="paragraph" w:customStyle="1" w:styleId="DocID"><w:name w:val="DocID"/>
  <w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="16"/></w:rPr></w:style>
<!-- revised -->
<w:style w:type="character" w:customStyle="1" w:styleId="DocID"><w:name w:val="DocID"/>
  <w:rPr><w:b w:val="0"/><w:sz w:val="16"/></w:rPr></w:style>
```

The [three comparison outputs supplied for #1187](https://github.com/JSv4/Docxodus/issues/1187#issuecomment-6415309420)
were produced with default options in Word for Mac, reported as approximately Word 2016. Their metadata
says `Microsoft Macintosh Word`, `AppVersion=15.0000`; the exact build and Formatting option were not
recorded. These observations describe those files, rather than every Word version or comparison setting.

| Revised document's `DocID` | Word's compare | Docxodus before #1185 | Docxodus after #1185 |
|---|---|---|---|
| Nothing in the revised document uses it | Original's paragraph style only in the #1184 real pair; **revised character style only** in the #1187 Mac control | Both styles under `DocID`: invalid package | Original's style only |
| An inserted run uses it (`w:rStyle="DocID"`) | Revised character style only, with `styleId=DocID`, `name=DocID`; inserted run still names `DocID` | Both styles under `DocID`: invalid package | Original's paragraph style keeps `DocID`; revised character style is copied as `DocID1` and the run names `DocID1` |
| An inserted paragraph uses it; the original's `DocID` is a character style | Revised paragraph style only, with `styleId=DocID`, `name=DocID`; inserted paragraph still names `DocID` | Both styles under `DocID`: invalid package | Original's character style keeps `DocID`; revised paragraph style is copied as `DocID1` and the paragraph names `DocID1` |

LibreOffice's own document comparison was not tried for these pairs.

**Analysis.** Docxodus keeps the original's styles part and copies in the revised styles it
lacks. The copy matched a revised style to an original one by type *and* id, so a revised character style
found no match and was copied in beside the paragraph style under the same id. Keying on the id alone
fixes the validity, but both source styles can still be needed. Dropping the revised style would point
the inserted run at a paragraph style, which a character reference cannot use, so the formatting would
be lost. Docxodus copies it under the first free id
`DocID1`, `DocID2`, …, and rewrites the references that mean it. Which style a reference means is read
off the reference: `w:rStyle` names a character style, `w:pStyle` a paragraph style, `w:tblStyle` a table
style, `w:numStyleLink`/`w:styleLink` a numbering style, `w:basedOn`/`w:next` a style of their own style's
type, and `w:link` the paragraph/character partner. A `w:style` without `w:type` is a paragraph style, and
the matching treats it as one.

A reference of the revised type to the colliding id cannot validly mean the original's style, but the
original can still hold one: a run whose `w:rStyle` names the original's paragraph style is invalid, yet
occurs. So the rewrite skips content only the original holds, where rejecting must give back the
original's reference unchanged: anything inside `w:del` or `w:moveFrom`, the archived old properties
inside `w:rPrChange`/`w:pPrChange`/…, the properties of a paragraph whose mark is deleted, and the
properties of a table whose every row is deleted. In the styles part and the numbering part only the
definitions copied from the revised document are rewritten. Content both documents share does follow the
revised style, which is what the revised document meant by it: an unchanged run is written from the
revised side, and a header or footer story the revision leaves unchanged is carried over from the
original but reads the same in both documents, so its references are rewritten as well. The copied style keeps its `w:name`, so the package carries two
styles with the display name `DocID`; the validator accepts that. Retaining both source names is
intentional; the style ids distinguish their references.

"Uses" counts only what enters the output: the stories that carry the revised content (body, headers,
footers, footnotes, endnotes, comments), outside original-only content; the revised styles that are
copied in (a `w:basedOn` or `w:link` from a revised style the output replaces with the original's does
not count); and the revised numbering definitions those stories or copied styles reach through
`w:numId` (`w:lvl/w:pStyle`). In a consolidation the output's body already holds every reviewer's
changes when a reviewer's styles are copied, so a reference there cannot say which reviewer it came
from; each reviewer's own package decides whether its style is used. That reviewer's numbering and
comment definitions are merged after its styles, and the rewrite is applied to them as they arrive. When
two reviewers both use a colliding style, the first reviewer's copy serves both, as for a same-type id.

**Code.** `Internal/StyleIdCollisions.cs` (collision test, use test, fresh ids, and the reference
rewrite `StyleIdRemap`), called from `IrMarkupRenderer.TrackStyleDefinitionChanges` (two-way compare) and
`PackageMerge.CopyMissingStylesFromOneDocToAnother` (consolidate). Tests: `DocxDiffCrossTypeStyleIdTests`.

**Decision for #1187.** The supplied Mac results neither rename nor merge the styles: they discard the
original style and retain the revised style's id and display name unchanged, even in the unused control.
In the referenced pair the original 8pt paragraph style is gone; in the reverse pair the original
bold/red character style is gone. Neither output archives the removed reference in a `w:pPrChange` or
`w:rPrChange`, so rejecting cannot restore that original formatting. The earlier #1184 observation is
therefore specific to its real document pair, not a universal rule about Word.

Docxodus deliberately keeps its existing policy rather than reproducing this formatting loss. Both
referenced styles survive under distinct ids, and formatting changes remain reversible. The two
referenced fixture shapes are checked through `DocxCompare.Compare`: accept and reject must match each
source's text, bold, colour and size after resolving style inheritance, with unique ids and no new
validation errors. No comparison implementation change was needed.

### Plain text inside a move-destination range (issue #1340)

A run between `w:moveToRangeStart` and `w:moveToRangeEnd`, outside any `w:moveTo`, is ordinary content. Word for the web preserves it when rejecting the move:

```xml
<w:moveToRangeStart w:id="3" w:name="move1" w:author="Alex" w:date="2026-05-01T09:00:00Z"/>
<w:moveTo w:id="4" w:author="Alex" w:date="2026-05-01T09:00:00Z"><w:r><w:t>amount</w:t></w:r></w:moveTo>
<w:r><w:t>XY</w:t></w:r>
<w:moveToRangeEnd w:id="3"/>
```

| Reader | Reject all |
|---|---|
| Word for the web, captured 2026-10-24 | keeps `XY`; removes only the moved destination text |
| LibreOffice | not adjudicated for these exact source inputs |
| `RevisionProcessor.RejectRevisions` before #1340 | drops `XY` |
| `RevisionProcessor.RejectRevisions` after #1340 | keeps `XY` |
| `DocxSession.RejectAllRevisions` | keeps `XY` |

Word also accepts and rejects two disjoint destination ranges with the same `w:name`, including a
range spanning paragraphs. The session now lists them as one supported move and resolves all its
ranges together. Each range still needs its own paired `w:id`, coherent author/date ownership and
ordered, disjoint source and destination content. Reused IDs, overlapping same-named ranges and
crossed ranges remain refused. A named range links explicitly marked moved content; its position
alone does not make ordinary content part of the move (ECMA-376 §17.13.5.24 and §17.13.5.28).

The processor inspects live content before removing moved wrappers. Range cleanup preserves
untracked runs and non-run content, including bookmarks, fields and originally empty content controls,
with their containers and properties. Controls emptied by the move and comment ranges whose moved
reference disappears do not preserve a second copy. Later passes still resolve deleted paragraph marks
and deleted content controls. This applies to
accepting source ranges and rejecting destination ranges.

`MoveRangeWordOracleTests` pins six original literal-OOXML source documents in
`TestFiles/RP/MoveRangeMembership/`: an unsplit control, plain text after and between moved wrappers,
a split destination, plain text outside the range, and a cross-paragraph split. Their manifest records
source hashes and native Word accept/reject DOCX, document XML and PDF hashes. All source packages
validate with Open XML SDK 3.5.1 (`Microsoft365`). The 36 endpoint checks exercise the stateless
processor, session bulk resolution and selective resolution against Word's actual paragraph lists;
all resulting packages must validate and contain no move markers. Native Word output is retained
unchanged in the evidence pack; some Word downloads contain unrelated settings-order findings, so
those outputs are evidence rather than validator-clean fixture inputs. Word exposed no build number.

Editing strictly inside a `w:moveTo` remains refused by `ReplaceTextAtSpan` and `SplitParagraph`:
coordinating changes to both moved sides is a separate editing operation. Text inserted at a same-
paragraph move edge still goes past its range marker. Existing ordinary content inside a range,
including text at an edge in another paragraph, now survives resolution by both APIs.

---

## Strict conformance: the same markup under different namespace URIs (issue #1268)

ISO/IEC 29500 Strict packages, which Word writes as "Strict Open XML Document", spell every
WordprocessingML name and relationship type in the `http://purl.oclc.org/ooxml/` family instead of
the transitional `http://schemas.openxmlformats.org/` one. For this reproducer the markup is otherwise
identical:

```xml
<w:document xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"><w:body><w:p>
  <w:del w:id="31" w:author="Review Example"><w:r><w:delText>Amber prism</w:delText></w:r></w:del>
  <w:ins w:id="32" w:author="Review Example"><w:r><w:t>Indigo prism</w:t></w:r></w:ins>
</w:p></w:body></w:document>
```

with `_rels/.rels` naming the main part through
`http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument`.

| Consumer | Strict input |
|---|---|
| Open XML SDK 3.5.1 validator (Office 2021) | 0 findings, as for the transitional twin |
| LibreOffice | renders the revisions as it does for the transitional twin |
| Word | converts Strict to transitional on open (not re-run for this reproducer) |
| Docxodus before #1268 | `DocxDiff`/`DocxCompare` and `HtmlConversionOps` converted; every other entry point matched transitional `W.*` names against Strict XML and found nothing — accept/reject left the markup, `HasTrackedRevisions` was `false`, `WmlToHtmlConverter.ConvertToHtml(WmlDocument)` threw because it found no `w:document` |
| Docxodus after | identical results to the transitional twin |

**Why.** The XDocument pipelines compare element names, and an element in a different namespace is a
different name. The Open XML SDK converts the namespaces when the package is written back, which is why accept/reject
returned a transitional package that still held the unresolved `w:ins`/`w:del`: the translation came
after the revision transforms had already skipped them.

**Code.** `StrictOoxmlNormalizer` rewrites the URIs. Code that opens a `WmlDocument` to read or change
its content takes the bytes from `DocxodusDocument.GetTransitionalBytes()`: the
`OpenXmlMemoryStreamDocument` constructor, `DocumentBuilder` (which converts each Strict source once on
entry), `DocumentStructureAnalyzer`, `MetricsGetter`, `OpenContractExporter`, `WmlDocument.MainDocumentPart`
and the rest. Code that reports on the package itself (hashes, manifests, deliverable verification)
keeps reading `DocumentByteArray`. `DocxSession` opens its live package from the conversion, keeps the
caller's exact bytes as the opening package (delivery source, verification baseline). `SemanticDiff`
runs its bounded package inspection on the bytes as given, then compares a Strict side's conversion,
so the semantic changes and delivery receipt of a session opened on Strict bytes show only its edits.
The Strict test reads only the main part's root start tag, so a transitional package costs one small
read. A damaged zip entry is copied through unconverted, and a package the conversion cannot read at
all is opened as it is, so damage fails where and how it did before the conversion existed. The `WordprocessingDocument` overloads receive a package the
caller already opened and are not converted. Tests: `StrictOoxmlRevisionTests`.

**Known limitation: value spaces.** `StrictOoxmlNormalizer` rewrites namespace and relationship URIs
only. Strict also narrows some value spaces, and those values pass through unchanged: `w:jc` is
`start`/`end` rather than `left`/`right`, `w:ind` uses `w:start`/`w:end`, and percentage widths are
written as `"50%"`. Docxodus already reads the `w:ind` and percent-suffixed spellings (see the sections
on each), but `WmlToHtmlConverter.CreateStyleFromJc` maps only `left`, `right`, `center` and `both`,
so a Strict `start`/`end` paragraph alignment is not rendered.


## Literal tabs inside text nodes (#1353)

The issue's Word example treats `<w:t xml:space="preserve">&#9;</w:t>` like `<w:tab/>`:
its position is the paragraph's next declared stop. A common docx-js shape has an empty run,
a tab-only text run, then a header label. The converter previously emitted the literal character
as browser whitespace, so a right-aligned label stayed near the left margin.

The converter now expands literal tabs in live and deleted text before the normal formatting
and tab-width pipeline. Body text, headers and footers use the same path. Each text fragment
keeps its attributes, formatting and revision owner; leading, consecutive and trailing tabs stay
in order. Field instructions retain literal tab whitespace and do not acquire layout tabs.
Native regressions compare complete visible text and tab markup with explicit-tab controls;
worker/pagination and export regressions check the physical right edge and PDF text/drawing facts.
Browser verification also found explicit tabs inside visible deletion wrappers using a full-width
spacer without accounting for the deleted text. The label wrapped back to the left margin. Revision
content now reuses the shared browser-aligned tab renderer, measuring visible deleted text and
skipping hidden deletion/move-source content. Guards keep field boundaries and revision ownership.
This implements the reported equivalence; no new Word capture is claimed.

## Floating shape clearance and stacking (#1346)

A 144 × 72pt `wps:wsp` rectangle, or the same rectangle inside `wpg:wgp`, can be
placed between two paragraphs with `wp:anchor`, column-relative horizontal offset 0,
paragraph-relative vertical offset 0, and `wp:wrapTopAndBottom`. The drawing is outside
its paragraph's text line. Dropping its SVG also drops its wrapping exclusion; treating
it as an inline SVG makes `wrapNone` consume height. The complete valid packages are
constructed by [the floating-shape fixture](../npm/tests/docx-floating-shape-fixture.ts).

The converter now separates SVG ink from text clearance. Top/bottom drawings reserve
the union of their bottom edges in the anchor paragraph. The paginator places the ink
and assigns drawing layers on the physical page. An explicit page stacking context is
necessary in standalone export: its cleanup removes `contain` and `will-change`, and
without another context a negative `z-index` drawing paints below the opaque paper.
The export regressions inspect PDF drawing/text operators as well as HTML geometry.

The local comparison used LibreOffice 25.8.7.3 and source-bound Chromium PDF exports.
Office Online was unavailable; these observations do not establish Word parity.

| Input | Word | LibreOffice | Docxodus |
| --- | --- | --- | --- |
| Paragraph-relative rectangle/group, top/bottom wrapping | Not captured | Keeps ink and clearance | Keeps ink and clearance |
| `behindDoc="1"`, no wrapping, with body text | Not captured | Ink behind text | Ink behind text and above paper |
| Teal `relativeHeight="4294967295"` followed by orange `relativeHeight="2147483648"` | Not captured | Orange above teal | Teal above orange |
| Filled rectangle with a triangular tight/through wrap polygon | Not captured | Uses rectangular clearance | Uses the supplied contour |
| Floating hexagon / simple inline VML rectangle | Not captured | Draws the geometry | Draws the geometry |

The SDK exposes `relativeHeight` as an unsigned 32-bit value
([Microsoft API reference](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.wordprocessing.anchor.relativeheight?view=openxml-3.0.1)).
Docxodus maps numeric order to dense CSS ranks rather than clamping it to CSS's signed
range. The high-value ordering and contour differences above need a Word capture
before either renderer can be treated as a Word reference for those variants. The
benchmark records them rather than changing the source values to match LibreOffice.

Hexagon geometry follows the ECMA preset equations (`adj=25000`, `vf=115470` by default),
including the shorter source side and subsequent group transform. The definitions were checked
against [LibreOffice's published copy at 5044a9a](https://github.com/LibreOffice/core/blob/5044a9a5719bba9402184438e5d66586de01d20f/oox/source/drawingml/customshapes/presetShapeDefinitions.xml).
Simple VML rectangles follow the declared fill/stroke flags and colours. Their default outline is
one pixel ([Microsoft VML reference](https://learn.microsoft.com/en-us/windows/win32/vml/msdn-online-vml-strokeweight-attribute));
unitless stroke weights are EMUs. Drawing these primitives avoids injecting a placeholder label
into the PDF text inventory. Other geometry, unsupported formulas and positioned VML retain the
ordinary unsupported-content policy.


## Tables: zero-width fixed cells keep a measurable sliver (issue #1201)

A fixed-layout table with `gridCol w:w="0"`, `tcW w:w="0" w:type="dxa"`, and zero
cell margins does not have an entirely invisible first cell in Word for the web. Native PDF
captures put the following cell about **1pt** after the body edge. With two zero grid columns
spanned by one cell, the borderless offset is still 1pt; two separate zero-width cells give
2pt. Explicit 12pt side padding remains additional to that content width. Default margins and
borders have further geometry differences; these observations do not resolve their general rules.

The converter gives this combination 1pt of content width and clips horizontal overflow.
It checks all grid columns the cell spans, its zero absolute preferred width, and fixed layout.
A missing grid width, a positive preferred width, and explicit autofit do not trigger the rule.
The shared cell-width calculation uses the same content width for tab placement. PageMap
registration retains measurable cell and paragraph fragments, including empty paragraphs;
there is no exception that forgives missing zero-width anchors.

The byte-exact native Word inputs, hashes, fonts, and capture geometry are recorded in
`TestFiles/ZeroWidthColumns/word-observations.json`. The original author fix is reused for
its grid-span lookup and horizontal clipping; the native captures correct its assumption
that the resulting column has exactly zero width. Tests: `ZeroWidthColumnTests` and
`npm/tests/zero-width-column.spec.ts`.

## Blocks taller than a page: exact rows, header rows, nested tables and border groups (#1202)

Native Word for the web exports of the exact inputs are recorded in
`TestFiles/OversizedPageBlocks/word-observations.json`. All twenty-three inputs have zero Open XML SDK
Microsoft365 findings. The browser regression checks compare the complete word multiset on every
page, including clipped words and duplicate occurrences, rather than only looking for selected markers.

Word places an oversized exact row alone on a new page and clips it at the paper edge. In the
300pt by 160pt twelve-line probes, Cedar01 through Cedar10 remain on page 2; Cedar11 and Cedar12 are
absent, and CedarAfter starts page 3. A footer overlaps the last visible row line. The same inventory
is observed when all twelve words are a single soft-wrapped text node. The standalone PDF writer
needs the existing whole-text clip expansion at this outer boundary too, or Skia drops the last
partly visible word. Text wholly outside the paper is made invisible without changing line boxes
or source anchors, so expanded clip bounds do not emit it on a later physical PDF page.

The paginator identifies the particular oversized exact row responsible for a box's overflow,
including its table-column boxes. Visible source anchors remain required. PageMap omits a wholly
clipped source only after checking its positive rectangle against the actual clipping edge. The
body's horizontal overflow remains visible; the page box clips at the paper. A real footnote band
keeps its protected body boundary.

A table's first fitting fragment participates in normal placement. This keeps OuterHead on page 1
in the nested-table case instead of moving the whole table. Word's oversized exact header cases,
including independent 200pt and 400pt controls, place the two body rows on page 3 and the following
paragraph on page 4. A conditional page break after those tables preserves that observation. A
400pt first row without the header flag keeps the following paragraph on page 3.

Bordered paragraphs and their shared-border groups split between their paragraph blocks. The
painted top/bottom sides are omitted at an internal continuation boundary, retaining the existing
physical border-width painter. Row spans and footnote-bearing table shapes retain their established
safety checks.

These observations establish page placement and visible text for the registered cases. They are
not a full Word raster score; native PDF bytes and font identities are retained with the captures.

A body ending in a table still has Word's required final paragraph, even if the package has no
serialized paragraph after the table. Its normal paragraph style determines whether it fits or
creates a blank final page. Native controls at 110pt and 130pt row heights, plus a 24pt default
line-height variant, distinguish this from a one-twip structural spacer. The converter inserts
this paragraph before formatting assembly as a view-only block with no source identity; explicit
final paragraphs remain unchanged. Tests retain every physical Word page, including blank ones,
and compare PDF crop boxes and exact per-page word inventories. Microsoft's
[blank-page guidance](https://support.microsoft.com/en-us/word/delete-a-page-in-word) describes
the required end paragraph and its ability to cause a blank last page.
