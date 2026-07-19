# OOXML Corner Cases

This document tracks edge cases and quirks in Open XML document processing where Word's behavior differs from a strict interpretation of the specification, or where the specification is ambiguous.

## Table of Contents

1. [Numbering and Lists](#numbering-and-lists)
   - [Legal Numbering with Multi-Level Format Strings](#legal-numbering-with-multi-level-format-strings)
2. [Footnotes](#footnotes)
   - [Footnote Count Discrepancy in Legal Templates](#footnote-count-discrepancy-in-legal-templates)
3. [Contributing](#contributing)

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

### The Word-parity alternative — `PreserveInputRevisions` and the one-sided Reject All

Word's own Compare does neither the default's leak nor the flag's flatten: it **preserves** the inputs'
pre-existing tracked revisions in the compare output verbatim (original author/date markup intact) while the
text diff is computed over the accepted view. Verified against Word-oracle outputs: an input with 176
revisions by another author keeps them in Word's compare result alongside the fresh compare revisions.
`DocxDiffSettings.PreserveInputRevisions` reproduces this (equal blocks + whole-block inserts, in the body
and in footnote/endnote bodies, in v1; it WINS
over `PreAcceptInputRevisions` when both are set, and the `DocxCompare` engine-selector path enables it).

The Word behavior worth pinning here: **Reject All on such an output does NOT restore the left document.**
Rejecting a preserved foreign `w:del` RESTORES its deleted text (text the left side never showed), and
rejecting a preserved foreign `w:ins` removes text the accepted view carried. Word's Compare output behaves
identically under Reject All — the one-sided round trip (`accept ≡ right` holds, `reject ≠ left` where
foreign markup exists) is inherent to preserving input revisions, not a Docxodus defect. Do not "fix" it.

### Relevant code

- `Docxodus/DocxDiff.cs` — `DocxDiffSettings.PreAcceptInputRevisions` + the `PreAccept(...)` pre-pass wired into
  all seven entry points; `IrMarkupRenderer.Render` clones the output on the LEFT package (the carry-over source);
  `DocxDiffSettings.PreserveInputRevisions` (the Word-parity opt-in, precedence over the pre-accept).
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

## Contributing

When adding new corner cases to this document:

1. **Provide a minimal reproducer**: Include the relevant XML snippets and a description of how to reproduce
2. **Document all renderers**: Test in Word, LibreOffice, and Docxodus
3. **Reference the spec**: Link to relevant ECMA-376 sections
4. **Identify the code**: Point to the specific Docxodus files/functions involved
5. **Propose a fix**: If possible, outline how the issue might be resolved
