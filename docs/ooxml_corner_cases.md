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

## Contributing

When adding new corner cases to this document:

1. **Provide a minimal reproducer**: Include the relevant XML snippets and a description of how to reproduce
2. **Document all renderers**: Test in Word, LibreOffice, and Docxodus
3. **Reference the spec**: Link to relevant ECMA-376 sections
4. **Identify the code**: Point to the specific Docxodus files/functions involved
5. **Propose a fix**: If possible, outline how the issue might be resolved
