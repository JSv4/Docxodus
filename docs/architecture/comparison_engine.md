# DOCX Comparison and HTML Rendering Engine

This document describes the architecture of the document comparison and tracked changes rendering system in Docxodus, covering how two Word documents are compared to produce tracked changes, and how those changes can be rendered to HTML.

## Overview

The comparison engine consists of two main components:

1. **WmlComparer** (`WmlComparer.cs`) - Compares two DOCX documents and produces a new document with tracked changes (revisions) marking insertions and deletions
2. **WmlToHtmlConverter** (`WmlToHtmlConverter.cs`) - Converts a DOCX document to HTML, optionally preserving and rendering tracked changes visually

These components can be used independently or together in a pipeline:

```
┌─────────────┐     ┌─────────────┐     ┌─────────────────────┐     ┌────────────┐
│  Original   │     │  Modified   │     │  Document with      │     │   HTML     │
│  Document   │────►│  Document   │────►│  Tracked Changes    │────►│  Output    │
│  (source1)  │     │  (source2)  │     │  (WmlComparer)      │     │            │
└─────────────┘     └─────────────┘     └─────────────────────┘     └────────────┘
                           │                      │                       │
                     WmlComparer.Compare()   WmlToHtmlConverter.ConvertToHtml()
                                                  │
                                          RenderTrackedChanges=true
```

## Part 1: Document Comparison (WmlComparer)

### High-Level Algorithm

The comparison algorithm uses a **Longest Common Subsequence (LCS)** approach at multiple levels:

1. **Preprocessing**: Normalize both documents
2. **Block-level correlation**: Find matching paragraphs/tables using hash codes
3. **Content-level comparison**: Compare text within correlated blocks using LCS
4. **Reconstruction**: Rebuild the document with revision markup

### Step 1: Preprocessing

Both documents undergo preprocessing to enable accurate comparison:

```csharp
source1 = PreProcessMarkup(source1, startingId + 1000);
source2 = PreProcessMarkup(source2, startingId + 2000);
```

**Key preprocessing steps:**

1. **Unid Assignment**: Every element gets a unique identifier (`pt:Unid` attribute)
   - Enables reconstruction of the XML tree after comparison
   - Preserved through revision acceptance/rejection where possible

2. **Markup Simplification**: Remove elements that don't affect content comparison
   - Comments, content controls, bookmarks, proofing marks, etc.

3. **Footnote/Endnote ID Remapping**: Ensure unique IDs across both documents
   - Prevents ID collisions when merging content

### Step 2: Block-Level Correlation

The algorithm establishes correlation between paragraphs by hash:

```
Source 1 (accept revisions)           Source 2 (reject revisions)
┌─────────────────────────────┐      ┌─────────────────────────────┐
│ Para 1: hash=ABC123         │◄────►│ Para 1: hash=ABC123         │  ← Correlated
│ Para 2: hash=DEF456         │      │ Para 2: hash=XYZ789         │  ← Different
│ Para 3: hash=GHI789         │◄────►│ Para 3: hash=GHI789         │  ← Correlated
└─────────────────────────────┘      └─────────────────────────────┘
```

**Process:**

1. Accept revisions in source1, calculate hash for each block-level element
2. Reject revisions in source2, calculate hash for each block-level element
3. Store hashes back in original documents using Unid correlation
4. Find longest common subsequence of matching block-level elements

### Step 3: Content Comparison

For each pair of correlated (or uncorrelated) blocks, the algorithm:

1. **Creates Comparison Unit Atoms**: Breaks content into atomic units
   - Each word, punctuation, or special element becomes an atom
   - Atoms contain content, formatting, and ancestor information

2. **Builds Comparison Unit Lists**: Groups atoms into comparison units
   - Paragraphs, table rows, cells, etc.

3. **Runs LCS on atoms**: Determines equal/inserted/deleted status

```csharp
var cal1 = CreateComparisonUnitAtomList(wDoc1.MainDocumentPart, body1, settings);
var cal2 = CreateComparisonUnitAtomList(wDoc2.MainDocumentPart, body2, settings);

var cus1 = GetComparisonUnitList(cal1, settings);
var cus2 = GetComparisonUnitList(cal2, settings);

var correlatedSequence = Lcs(cus1, cus2, settings);
```

### Comparison Unit Structure

```
ComparisonUnitGroup
├── ComparisonUnit (paragraph)
│   ├── ComparisonUnitAtom (word: "Hello")
│   ├── ComparisonUnitAtom (space)
│   └── ComparisonUnitAtom (word: "World")
└── ComparisonUnit (paragraph)
    └── ...
```

Each atom tracks:
- **Content**: The actual text or element
- **AncestorUnids**: Path from root to this element (for reconstruction)
- **CorrelationStatus**: `Equal`, `Inserted`, `Deleted`, or `Unknown`

### Step 4: Document Reconstruction

After LCS determines what's equal/inserted/deleted:

1. **Flatten to atom list**: Get ordered list with status markers
2. **Assemble ancestor Unids**: Ensure proper XML tree reconstruction
3. **Produce WordprocessingML document**: Build new document with revision markup

```csharp
var listOfComparisonUnitAtoms = FlattenToComparisonUnitAtomList(correlatedSequence, settings);
AssembleAncestorUnidsInOrderToRebuildXmlTreeProperly(listOfComparisonUnitAtoms);
var result = ProduceDocumentWithRevisions(listOfComparisonUnitAtoms, settings);
```

### OOXML Revision Elements

The comparer generates standard OOXML revision elements:

| Element | Description |
|---------|-------------|
| `w:ins` | Inserted content - wraps runs that were added |
| `w:del` | Deleted content - wraps runs that were removed |
| `w:moveFrom` | Source location of moved content |
| `w:moveTo` | Destination location of moved content |
| `w:rPrChange` | Run property (formatting) change |
| `w:pPrChange` | Paragraph property change |
| `w:trPr/w:ins` | Table row was inserted |
| `w:trPr/w:del` | Table row was deleted |
| `w:tcPr/w:cellIns` | Table cell was inserted |
| `w:tcPr/w:cellDel` | Table cell was deleted |

**Revision attributes:**
- `w:id` - Unique revision identifier
- `w:author` - Author name (from `WmlComparerSettings.AuthorForRevisions`)
- `w:date` - ISO 8601 timestamp

### WmlComparerSettings

```csharp
public class WmlComparerSettings
{
    // Author name for generated revisions
    public string AuthorForRevisions = "Open-Xml-PowerTools";

    // Timestamp for revisions (ISO 8601 format)
    public string DateTimeForRevisions = DateTime.Now.ToString("o");

    // Threshold for word-level vs character-level comparison
    // Lower = more detailed comparison within changed regions
    public double DetailThreshold = 0.15;

    // Case-insensitive comparison
    public bool CaseInsensitive = false;

    // Treat breaking and non-breaking spaces as equivalent
    public bool ConflateBreakingAndNonbreakingSpaces = true;

    // Word boundary characters
    public char[] WordSeparators = new[] { ' ', '-', ')', '(', ';', ',', ... };

    // Culture for string comparison
    public CultureInfo CultureInfo = null;
}
```

---

## Part 2: Tracked Changes HTML Rendering

The `WmlToHtmlConverter` can render tracked changes visually in HTML output, rather than accepting them before conversion.

### Traditional Behavior (Default)

By default, revisions are accepted before conversion:

```
┌─────────────────────┐     ┌─────────────────────┐     ┌─────────────────────┐
│   DOCX with         │     │   DOCX after        │     │   HTML              │
│   tracked changes   │────►│   AcceptRevisions() │────►│   (clean output)    │
│                     │     │                     │     │                     │
│   "Hello World"     │     │   "Hello World"     │     │   Hello World       │
│   ──────   ─────    │     │                     │     │                     │
│   del      ins      │     │                     │     │                     │
└─────────────────────┘     └─────────────────────┘     └─────────────────────┘
```

### Tracked Changes Rendering (New)

With `RenderTrackedChanges = true`:

```
┌─────────────────────┐     ┌─────────────────────────────────────────────────┐
│   DOCX with         │     │   HTML with revision styling                    │
│   tracked changes   │────►│                                                 │
│                     │     │   <del class="rev-del">Hello</del>              │
│   "Hello World"     │     │   <ins class="rev-ins">World</ins>              │
│   ──────   ─────    │     │                                                 │
│   del      ins      │     │   Styled with strikethrough/underline + colors  │
└─────────────────────┘     └─────────────────────────────────────────────────┘
```

### HTML Output Format

#### Insertions (`w:ins`)

```html
<ins class="rev-ins" data-author="John Doe" data-date="2024-01-15T10:30:00Z">
  <span class="pt-000001">inserted text</span>
</ins>
```

#### Deletions (`w:del`)

```html
<del class="rev-del" data-author="Jane Smith" data-date="2024-01-14T09:00:00Z">
  <span class="pt-000002">deleted text</span>
</del>
```

#### Move Operations

When `RenderMoveOperations = true`:

```html
<!-- Move source (where content came from) -->
<del class="rev-move-from" data-move-id="1" data-author="..." data-date="...">
  <span>moved text</span>
</del>

<!-- Move destination (where content went to) -->
<ins class="rev-move-to" data-move-id="1" data-author="..." data-date="...">
  <span>moved text</span>
</ins>
```

#### Table Row/Cell Revisions

```html
<tr class="rev-row-ins" data-author="..." data-date="...">
  <td>new row content</td>
</tr>

<tr class="rev-row-del" data-author="..." data-date="...">
  <td>deleted row content</td>
</tr>

<td class="rev-cell-ins">new cell</td>
<td class="rev-cell-del">deleted cell</td>
```

#### Format Changes (`w:rPrChange`)

```html
<span class="rev-format-change"
      title="Format changed: Bold added, Italic removed"
      data-author="..."
      data-date="...">
  formatted text
</span>
```

### Generated CSS

The converter generates CSS for revision styling:

```css
/* Tracked Changes CSS */

/* Insertions */
ins.rev-ins {
    text-decoration: underline;
    color: #006400;       /* dark green */
    background-color: #e6ffe6;
}

/* Deletions */
del.rev-del {
    text-decoration: line-through;
    color: #8b0000;       /* dark red */
    background-color: #ffe6e6;
}

/* Move source */
del.rev-move-from {
    text-decoration: line-through;
    color: #4b0082;       /* indigo */
    background-color: #f0e6ff;
}

/* Move destination */
ins.rev-move-to {
    text-decoration: underline;
    color: #4b0082;
    background-color: #e6f0ff;
}

/* Table row changes */
tr.rev-row-ins {
    background-color: #e6ffe6;
}

tr.rev-row-del {
    background-color: #ffe6e6;
    text-decoration: line-through;
}

/* Cell changes */
td.rev-cell-ins {
    background-color: #e6ffe6;
}

td.rev-cell-del {
    background-color: #ffe6e6;
    text-decoration: line-through;
}

/* Paragraph mark indicator */
span.rev-para-mark {
    color: #888;
    font-size: 0.8em;
    vertical-align: super;
}

/* Format changes */
span.rev-format-change {
    border-bottom: 2px dotted #ffa500;
}

/* Author-specific colors (if AuthorColors provided) */
[data-author="John Doe"] {
    border-left: 3px solid #0066cc;
}
```

### WmlToHtmlConverterSettings

```csharp
public class WmlToHtmlConverterSettings
{
    // === Tracked Changes Settings ===

    /// <summary>
    /// If true, render tracked changes visually in HTML output.
    /// If false (default), accept all revisions before conversion.
    /// </summary>
    public bool RenderTrackedChanges = false;

    /// <summary>
    /// CSS class prefix for revision elements (default: "rev-")
    /// </summary>
    public string RevisionCssClassPrefix = "rev-";

    /// <summary>
    /// If true, include revision metadata (author, date) as data attributes
    /// </summary>
    public bool IncludeRevisionMetadata = true;

    /// <summary>
    /// If true, show deleted content with strikethrough (default: true)
    /// If false, hide deleted content entirely
    /// </summary>
    public bool ShowDeletedContent = true;

    /// <summary>
    /// Custom colors for different authors (author name -> CSS color)
    /// Generates CSS rules like [data-author="Name"] { border-left: 3px solid color; }
    /// </summary>
    public Dictionary<string, string> AuthorColors;

    /// <summary>
    /// If true, render move operations as separate from/to (default: true)
    /// If false, render moves as regular delete + insert
    /// </summary>
    public bool RenderMoveOperations = true;

    // === Additional Content Settings ===

    /// <summary>
    /// If true, render footnotes and endnotes at the end of the HTML document.
    /// If false (default), footnotes and endnotes are stripped.
    /// </summary>
    public bool RenderFootnotesAndEndnotes = false;

    /// <summary>
    /// If true, render headers and footers in the HTML document.
    /// If false (default), headers and footers are not rendered.
    /// </summary>
    public bool RenderHeadersAndFooters = false;
}
```

### Processing Pipeline

```
┌─────────────────────────────────────────────────────────────────────────────────┐
│                        WmlToHtmlConverter Pipeline                              │
├─────────────────────────────────────────────────────────────────────────────────┤
│                                                                                 │
│  if (!RenderTrackedChanges)                                                     │
│      RevisionAccepter.AcceptRevisions(wordDoc);  ◄── Skip this step            │
│                                                                                 │
│  MarkupSimplifier.SimplifyMarkup(wordDoc, settings);                           │
│      ├── RemoveComments (configurable)                                         │
│      ├── RemoveContentControls                                                 │
│      ├── RemoveProof                                                           │
│      └── Keep revision elements if RenderTrackedChanges                        │
│                                                                                 │
│  FormattingAssembler.AssembleFormatting(wordDoc, settings);                    │
│      └── Resolve style inheritance, flatten formatting                         │
│                                                                                 │
│  ConvertToHtmlTransform(element)  ◄── Recursive transformation                 │
│      │                                                                          │
│      ├── w:ins  → ProcessInsertion() → <ins class="rev-ins">                   │
│      ├── w:del  → ProcessDeletion()  → <del class="rev-del">                   │
│      ├── w:moveFrom → ProcessMoveFrom() → <del class="rev-move-from">          │
│      ├── w:moveTo   → ProcessMoveTo()   → <ins class="rev-move-to">            │
│      ├── w:delText  → XText (deleted text content)                             │
│      ├── w:p        → ProcessParagraph() (handles pPr revisions)               │
│      ├── w:tr       → ProcessTableRow() (handles trPr revisions)               │
│      ├── w:tc       → ProcessTableCell() (handles tcPr revisions)              │
│      ├── w:r        → ProcessRun() (handles rPrChange)                         │
│      └── ... other elements                                                     │
│                                                                                 │
│  ReifyStylesAndClasses()                                                        │
│      ├── Generate CSS classes or inline styles                                 │
│      └── Add revision CSS if RenderTrackedChanges                              │
│                                                                                 │
└─────────────────────────────────────────────────────────────────────────────────┘
```

### Handler Implementation Details

#### ProcessInsertion

```csharp
private static object ProcessInsertion(WordprocessingDocument wordDoc,
    WmlToHtmlConverterSettings settings, XElement element, decimal currentMarginLeft)
{
    if (!settings.RenderTrackedChanges)
    {
        // Fallback: just process children normally
        return element.Elements()
            .Select(e => ConvertToHtmlTransform(wordDoc, settings, e, false, currentMarginLeft));
    }

    var ins = new XElement(Xhtml.ins);
    ins.Add(new XAttribute("class", settings.RevisionCssClassPrefix + "ins"));

    if (settings.IncludeRevisionMetadata)
    {
        var author = (string)element.Attribute(W.author);
        var date = (string)element.Attribute(W.date);
        if (author != null) ins.Add(new XAttribute("data-author", author));
        if (date != null) ins.Add(new XAttribute("data-date", date));
    }

    // Recursively process children
    ins.Add(element.Elements()
        .Select(e => ConvertToHtmlTransform(wordDoc, settings, e, false, currentMarginLeft)));

    return ins;
}
```

#### ProcessDeletion

Key difference: handles `w:delText` elements (deleted text uses `w:delText` instead of `w:t`):

```csharp
private static object ProcessDeletion(...)
{
    if (!settings.RenderTrackedChanges) return null;  // Remove entirely

    if (!settings.ShowDeletedContent)
    {
        // Return marker without content
        return new XElement(Xhtml.span,
            new XAttribute("class", "rev-del-marker"),
            new XAttribute("title", "Deleted content"));
    }

    var del = new XElement(Xhtml.del);
    // ... add class and metadata ...
    del.Add(element.Elements()
        .Select(e => ConvertToHtmlTransform(...)));
    return del;
}
```

The `w:delText` element handler:
```csharp
if (element.Name == W.delText && settings.RenderTrackedChanges)
{
    return new XText(element.Value);  // Extract the deleted text
}
```

---

## Usage Examples

### Compare Documents and Render to HTML

```csharp
using Docxodus;

// Load documents
var original = new WmlDocument("original.docx");
var modified = new WmlDocument("modified.docx");

// Compare
var comparerSettings = new WmlComparerSettings
{
    AuthorForRevisions = "Document Comparison",
    DetailThreshold = 0.15
};

var compared = WmlComparer.Compare(original, modified, comparerSettings);

// Convert to HTML with tracked changes visible
var htmlSettings = new WmlToHtmlConverterSettings
{
    PageTitle = "Document Comparison",
    RenderTrackedChanges = true,
    IncludeRevisionMetadata = true,
    ShowDeletedContent = true,
    RenderMoveOperations = true,
    AuthorColors = new Dictionary<string, string>
    {
        { "Document Comparison", "#0066cc" }
    }
};

var html = WmlToHtmlConverter.ConvertToHtml(compared, htmlSettings);
html.Save("comparison.html");
```

### Render Existing Tracked Changes

```csharp
// Load document that already has tracked changes
var docWithRevisions = new WmlDocument("document-with-revisions.docx");

var settings = new WmlToHtmlConverterSettings
{
    PageTitle = "Review Document",
    RenderTrackedChanges = true,
    RenderFootnotesAndEndnotes = true,
    RenderHeadersAndFooters = true,
    AuthorColors = new Dictionary<string, string>
    {
        { "John Doe", "#ff6600" },
        { "Jane Smith", "#0066ff" }
    }
};

var html = docWithRevisions.ConvertToHtml(settings);
```

---

## Architecture Diagram

```
┌────────────────────────────────────────────────────────────────────────────────┐
│                              Docxodus Library                                   │
├────────────────────────────────────────────────────────────────────────────────┤
│                                                                                 │
│  ┌──────────────────────┐        ┌──────────────────────┐                      │
│  │    WmlComparer       │        │  WmlToHtmlConverter  │                      │
│  │                      │        │                      │                      │
│  │  ┌────────────────┐  │        │  ┌────────────────┐  │                      │
│  │  │ PreProcess     │  │        │  │ SimplifyMarkup │  │                      │
│  │  │ Markup         │  │        │  │                │  │                      │
│  │  └───────┬────────┘  │        │  └───────┬────────┘  │                      │
│  │          │           │        │          │           │                      │
│  │  ┌───────▼────────┐  │        │  ┌───────▼────────┐  │                      │
│  │  │ Create         │  │        │  │ Formatting     │  │                      │
│  │  │ ComparisonUnits│  │        │  │ Assembler      │  │                      │
│  │  └───────┬────────┘  │        │  └───────┬────────┘  │                      │
│  │          │           │        │          │           │                      │
│  │  ┌───────▼────────┐  │        │  ┌───────▼────────┐  │                      │
│  │  │ LCS Algorithm  │  │        │  │ ConvertToHtml  │  │                      │
│  │  │                │  │        │  │ Transform      │  │                      │
│  │  └───────┬────────┘  │        │  │                │  │                      │
│  │          │           │        │  │ ├─ ProcessIns  │  │                      │
│  │  ┌───────▼────────┐  │        │  │ ├─ ProcessDel  │  │                      │
│  │  │ Produce        │  │        │  │ ├─ ProcessMove │  │                      │
│  │  │ Revisions      │  │        │  │ └─ etc.        │  │                      │
│  │  └───────┬────────┘  │        │  └───────┬────────┘  │                      │
│  │          │           │        │          │           │                      │
│  └──────────┼───────────┘        │  ┌───────▼────────┐  │                      │
│             │                    │  │ Generate CSS   │  │                      │
│             │                    │  │ (inc. revision │  │                      │
│             │                    │  │  styles)       │  │                      │
│             │                    │  └───────┬────────┘  │                      │
│             │                    │          │           │                      │
│             │                    └──────────┼───────────┘                      │
│             │                               │                                   │
│             ▼                               ▼                                   │
│  ┌──────────────────────┐        ┌──────────────────────┐                      │
│  │   WmlDocument        │        │   XElement (HTML)    │                      │
│  │   with revisions     │        │   with revision      │                      │
│  │   (w:ins, w:del,     │───────►│   markup (<ins>,     │                      │
│  │    w:moveFrom, etc.) │        │   <del>, etc.)       │                      │
│  └──────────────────────┘        └──────────────────────┘                      │
│                                                                                 │
└────────────────────────────────────────────────────────────────────────────────┘
```

---

## Test Coverage

The tracked changes HTML rendering is covered by tests HC001-HC013:

| Test | Description |
|------|-------------|
| HC001 | Basic document conversion |
| HC002 | Complex document conversion |
| HC003 | Insertion rendering (`w:ins` → `<ins>`) |
| HC004 | Deletion rendering (`w:del` → `<del>`) |
| HC005 | Tracked changes disabled (default behavior) |
| HC006 | Tracked changes CSS generation |
| HC007 | Footnote/endnote rendering |
| HC008 | Format change rendering (`w:rPrChange`) |
| HC009 | Header rendering |
| HC010 | Footer rendering |
| HC011 | Move operations (`w:moveFrom`/`w:moveTo`) |
| HC012 | Author-specific color styling |
| HC013 | All features enabled together |

---

## Related Files

| File | Description |
|------|-------------|
| `Docxodus/WmlComparer.cs` | Document comparison algorithm |
| `Docxodus/WmlToHtmlConverter.cs` | HTML conversion with tracked changes |
| `Docxodus/RevisionAccepter.cs` | Accept tracked revisions |
| `Docxodus/RevisionProcessor.cs` | Accept/reject revision processing |
| `Docxodus/PtOpenXmlUtil.cs` | XML utilities, namespace definitions |
| `Docxodus/TRACKED_CHANGES_HTML_DESIGN.md` | Design document for tracked changes feature |
| `Docxodus.Tests/HtmlConverterTests.cs` | HTML converter tests |
| `Docxodus.Tests/WcTests.cs` | WmlComparer tests |
