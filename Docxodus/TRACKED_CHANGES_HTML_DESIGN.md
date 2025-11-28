# Tracked Changes HTML Rendering - Design Document

## Executive Summary

This document outlines a design for extending `WmlToHtmlConverter` to support rendering tracked changes (revisions) in the HTML output, rather than silently accepting them before conversion.

## Current Behavior

The current `WmlToHtmlConverter.ConvertToHtml()` method (line 169) calls:

```csharp
RevisionAccepter.AcceptRevisions(wordDoc);
```

This **removes all tracked changes** before conversion, meaning:
- Inserted text is kept, but the insertion marking is lost
- Deleted text is completely removed
- Move operations are flattened
- Property changes (formatting, paragraph, table) are accepted

## Revision Elements in OOXML

### Content Revisions

| Element | Description | Content |
|---------|-------------|---------|
| `w:ins` | Inserted content | Contains `w:r` runs with `w:t` text |
| `w:del` | Deleted content | Contains `w:r` runs with `w:delText` text |
| `w:moveFrom` | Source of moved content | Contains runs (like `w:del`) |
| `w:moveTo` | Destination of moved content | Contains runs (like `w:ins`) |

### Paragraph/Run Mark Revisions

| Element | Location | Meaning |
|---------|----------|---------|
| `w:ins` | `w:pPr/w:rPr/w:ins` | Paragraph mark was inserted |
| `w:del` | `w:pPr/w:rPr/w:del` | Paragraph mark was deleted |

### Table Revisions

| Element | Location | Meaning |
|---------|----------|---------|
| `w:ins`/`w:del` | `w:trPr` | Row inserted/deleted |
| `w:cellIns` | `w:tcPr` | Cell inserted |
| `w:cellDel` | `w:tcPr` | Cell deleted |
| `w:cellMerge` | `w:tcPr` | Cell merged |

### Property Change Revisions

| Element | Description |
|---------|-------------|
| `w:rPrChange` | Run formatting change |
| `w:pPrChange` | Paragraph formatting change |
| `w:sectPrChange` | Section formatting change |
| `w:tblPrChange` | Table formatting change |
| `w:trPrChange` | Table row formatting change |
| `w:tcPrChange` | Table cell formatting change |
| `w:tblGridChange` | Table grid change |

### Common Attributes

All revision elements have:
- `w:id` - Unique revision ID
- `w:author` - Author name
- `w:date` - ISO 8601 timestamp

## Proposed Design

### 1. Settings Extension

Add new properties to `WmlToHtmlConverterSettings`:

```csharp
public class WmlToHtmlConverterSettings
{
    // ... existing properties ...

    /// <summary>
    /// If true, render tracked changes visually in HTML output.
    /// If false (default), accept all revisions before conversion.
    /// </summary>
    public bool RenderTrackedChanges;

    /// <summary>
    /// CSS class prefix for revision elements (default: "rev-")
    /// </summary>
    public string RevisionCssClassPrefix;

    /// <summary>
    /// If true, include revision metadata (author, date) as data attributes
    /// </summary>
    public bool IncludeRevisionMetadata;

    /// <summary>
    /// Custom colors for different authors (author name -> CSS color)
    /// </summary>
    public Dictionary<string, string> AuthorColors;

    /// <summary>
    /// If true, show deleted content with strikethrough (default: true)
    /// If false, hide deleted content but mark its location
    /// </summary>
    public bool ShowDeletedContent;

    /// <summary>
    /// If true, render move operations as separate from/to
    /// If false, render moves as delete + insert
    /// </summary>
    public bool RenderMoveOperations;
}
```

### 2. HTML Output Format

#### Insertions

```html
<ins class="rev-ins" data-author="John Doe" data-date="2024-01-15T10:30:00Z">
  <span>inserted text</span>
</ins>
```

#### Deletions

```html
<del class="rev-del" data-author="Jane Smith" data-date="2024-01-14T09:00:00Z">
  <span>deleted text</span>
</del>
```

#### Move Operations (if `RenderMoveOperations` is true)

```html
<!-- Move source -->
<del class="rev-move-from" data-move-id="move123" data-author="..." data-date="...">
  <span>moved text</span>
</del>

<!-- Move destination -->
<ins class="rev-move-to" data-move-id="move123" data-author="..." data-date="...">
  <span>moved text</span>
</ins>
```

#### Paragraph Mark Changes

```html
<!-- Deleted paragraph mark (paragraphs were merged) -->
<p class="rev-para-del" data-author="..." data-date="...">
  content...
  <span class="rev-para-mark-del" title="Paragraph mark deleted">¶</span>
</p>

<!-- Inserted paragraph (split from another) -->
<p class="rev-para-ins" data-author="..." data-date="...">
  content...
</p>
```

#### Table Row Changes

```html
<tr class="rev-row-ins" data-author="..." data-date="...">
  <td>...</td>
</tr>

<tr class="rev-row-del" data-author="..." data-date="...">
  <td>...</td>
</tr>
```

#### Formatting Changes

```html
<span class="rev-format-change"
      data-author="..."
      data-date="..."
      data-change-type="rPrChange"
      title="Format changed: Bold added">
  formatted text
</span>
```

### 3. Default CSS

Generate appropriate CSS in the `<style>` element:

```css
/* Insertions */
ins.rev-ins {
  text-decoration: underline;
  color: #006400; /* dark green */
  background-color: #e6ffe6;
}

/* Deletions */
del.rev-del {
  text-decoration: line-through;
  color: #8b0000; /* dark red */
  background-color: #ffe6e6;
}

/* Move source */
del.rev-move-from {
  text-decoration: line-through;
  color: #4b0082; /* indigo */
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

/* Paragraph mark indicator */
.rev-para-mark-del {
  color: #8b0000;
  font-size: 0.8em;
  vertical-align: super;
}

/* Format changes */
.rev-format-change {
  border-bottom: 2px dotted #ffa500;
}
```

### 4. Implementation Changes

#### 4.1. Entry Point Modification

```csharp
public static XElement ConvertToHtml(WordprocessingDocument wordDoc,
    WmlToHtmlConverterSettings htmlConverterSettings)
{
    // Only accept revisions if NOT rendering tracked changes
    if (!htmlConverterSettings.RenderTrackedChanges)
    {
        RevisionAccepter.AcceptRevisions(wordDoc);
    }

    // ... rest of existing code ...
}
```

#### 4.2. New Handler Methods

Add handlers in `ConvertToHtmlTransform()`:

```csharp
// Handle w:ins (inserted content)
if (element.Name == W.ins)
{
    return ProcessInsertion(wordDoc, settings, element, currentMarginLeft);
}

// Handle w:del (deleted content)
if (element.Name == W.del)
{
    return ProcessDeletion(wordDoc, settings, element, currentMarginLeft);
}

// Handle w:moveFrom
if (element.Name == W.moveFrom)
{
    return ProcessMoveFrom(wordDoc, settings, element, currentMarginLeft);
}

// Handle w:moveTo
if (element.Name == W.moveTo)
{
    return ProcessMoveTo(wordDoc, settings, element, currentMarginLeft);
}
```

#### 4.3. ProcessInsertion Implementation

```csharp
private static object ProcessInsertion(WordprocessingDocument wordDoc,
    WmlToHtmlConverterSettings settings, XElement element, decimal currentMarginLeft)
{
    if (!settings.RenderTrackedChanges)
    {
        // Fall through to process children normally
        return element.Elements()
            .Select(e => ConvertToHtmlTransform(wordDoc, settings, e, false, currentMarginLeft));
    }

    var ins = new XElement(Xhtml.ins);

    // Add CSS class
    var className = (settings.RevisionCssClassPrefix ?? "rev-") + "ins";
    ins.Add(new XAttribute("class", className));

    // Add metadata if requested
    if (settings.IncludeRevisionMetadata)
    {
        var author = (string)element.Attribute(W.author);
        var date = (string)element.Attribute(W.date);

        if (author != null)
            ins.Add(new XAttribute("data-author", author));
        if (date != null)
            ins.Add(new XAttribute("data-date", date));
    }

    // Process children
    ins.Add(element.Elements()
        .Select(e => ConvertToHtmlTransform(wordDoc, settings, e, false, currentMarginLeft)));

    return ins;
}
```

#### 4.4. ProcessDeletion Implementation

```csharp
private static object ProcessDeletion(WordprocessingDocument wordDoc,
    WmlToHtmlConverterSettings settings, XElement element, decimal currentMarginLeft)
{
    if (!settings.RenderTrackedChanges)
    {
        // When not rendering tracked changes, deletions are removed
        return null;
    }

    if (!settings.ShowDeletedContent)
    {
        // Show marker but not content
        return new XElement(Xhtml.span,
            new XAttribute("class", (settings.RevisionCssClassPrefix ?? "rev-") + "del-marker"),
            new XAttribute("title", "Deleted content"));
    }

    var del = new XElement(Xhtml.del);

    // Add CSS class
    var className = (settings.RevisionCssClassPrefix ?? "rev-") + "del";
    del.Add(new XAttribute("class", className));

    // Add metadata if requested
    if (settings.IncludeRevisionMetadata)
    {
        var author = (string)element.Attribute(W.author);
        var date = (string)element.Attribute(W.date);

        if (author != null)
            del.Add(new XAttribute("data-author", author));
        if (date != null)
            del.Add(new XAttribute("data-date", date));
    }

    // Process children - note: w:del contains w:delText instead of w:t
    del.Add(ProcessDeletedContent(wordDoc, settings, element, currentMarginLeft));

    return del;
}
```

#### 4.5. Handle w:delText

```csharp
// Transform every w:delText element to a text node (only when rendering revisions)
if (element.Name == W.delText && settings.RenderTrackedChanges)
{
    return new XText(element.Value);
}
```

#### 4.6. Paragraph Mark Revisions

Modify `ProcessParagraph` to check for `w:pPr/w:rPr/w:ins` or `w:del`:

```csharp
private static object ProcessParagraph(...)
{
    // ... existing code ...

    if (settings.RenderTrackedChanges)
    {
        var pPr = element.Element(W.pPr);
        var rPr = pPr?.Element(W.rPr);

        var paraIns = rPr?.Element(W.ins);
        var paraDel = rPr?.Element(W.del);

        if (paraIns != null)
        {
            // This paragraph mark was inserted (paragraph was split)
            style.Add("--rev-para-ins", "true");
            // Add appropriate class to paragraph element
        }

        if (paraDel != null)
        {
            // This paragraph mark was deleted (paragraphs were merged)
            // Add pilcrow marker at end of paragraph
        }
    }

    // ... rest of existing code ...
}
```

#### 4.7. Table Row Revisions

Modify `ProcessTableRow` to handle `w:trPr/w:ins` and `w:trPr/w:del`:

```csharp
private static object ProcessTableRow(...)
{
    var style = new Dictionary<string, string>();

    if (settings.RenderTrackedChanges)
    {
        var trPr = element.Element(W.trPr);
        var rowIns = trPr?.Element(W.ins);
        var rowDel = trPr?.Element(W.del);

        if (rowIns != null || rowDel != null)
        {
            // Add class to row
        }
    }

    // ... rest of existing code ...
}
```

### 5. CSS Generation

Add revision CSS to `ReifyStylesAndClasses`:

```csharp
private static void ReifyStylesAndClasses(...)
{
    // ... existing code ...

    if (htmlConverterSettings.RenderTrackedChanges)
    {
        sb.Append(GenerateRevisionCss(htmlConverterSettings));
    }

    // ... rest of existing code ...
}

private static string GenerateRevisionCss(WmlToHtmlConverterSettings settings)
{
    var prefix = settings.RevisionCssClassPrefix ?? "rev-";
    var sb = new StringBuilder();

    sb.AppendLine($"ins.{prefix}ins {{ text-decoration: underline; color: #006400; }}");
    sb.AppendLine($"del.{prefix}del {{ text-decoration: line-through; color: #8b0000; }}");
    // ... more CSS ...

    // Author-specific colors
    if (settings.AuthorColors != null)
    {
        foreach (var kvp in settings.AuthorColors)
        {
            var safeAuthor = EscapeCssSelector(kvp.Key);
            sb.AppendLine($"[data-author=\"{safeAuthor}\"] {{ border-left: 3px solid {kvp.Value}; }}");
        }
    }

    return sb.ToString();
}
```

### 6. Processing Pipeline Adjustments

When `RenderTrackedChanges` is enabled, several preprocessing steps need adjustment:

1. **Skip `RevisionAccepter.AcceptRevisions()`** - Already handled
2. **Adjust `SimplifyMarkupSettings`**:
   - Keep `RemoveComments = true` (or make configurable)
   - Keep `RemoveProof = true`
   - Ensure revision elements are NOT removed

3. **Adjust `FormattingAssembler`**:
   - Need to ensure it doesn't strip revision metadata
   - May need to pass through revision attributes

### 7. Edge Cases to Handle

1. **Nested revisions**: An insertion containing a deletion (rare but possible)
2. **Split revisions**: A single revision spanning multiple paragraphs
3. **Revisions in tables**: Rows, cells, and content all have different handling
4. **Revisions in headers/footers**: Need to process all document parts
5. **Revisions in footnotes/endnotes**: Similar to headers/footers
6. **Field codes with revisions**: `w:delInstrText` vs `w:instrText`
7. **Math with revisions**: `m:r` elements containing `w:ins`/`w:del`
8. **Content controls with revisions**: `customXmlInsRangeStart`, etc.

### 8. Testing Strategy

Create test cases for:

1. Simple text insertion
2. Simple text deletion
3. Multiple insertions by different authors
4. Move operations
5. Paragraph merge (deleted paragraph mark)
6. Paragraph split (inserted paragraph mark)
7. Table row insertion/deletion
8. Table cell insertion/deletion
9. Formatting changes
10. Nested revisions
11. Revisions in footnotes
12. Revisions in headers/footers
13. Complex documents with mixed revision types

### 9. Implementation Phases

#### Phase 1: Core Infrastructure ✅ COMPLETE
- [x] Add settings properties (RenderTrackedChanges, RevisionCssClassPrefix, etc.)
- [x] Modify entry point to conditionally skip revision acceptance
- [x] Add `w:ins` and `w:del` handlers for inline content
- [x] Handle `w:delText`
- [x] Generate basic CSS
- [x] Add unit tests (HC003-HC006)

#### Phase 2: Extended Content Types ✅ COMPLETE
- [x] Move operations (`w:moveFrom`/`w:moveTo`)
- [x] RenderMoveOperations setting
- [x] Paragraph mark revisions (`w:pPr/w:rPr/w:ins` and `w:del`)
- [x] Table row revisions (`w:trPr/w:ins` and `w:del`)
- [x] CSS for moves, paragraph marks, table rows

#### Phase 3: Advanced Features ✅ COMPLETE
- [x] Property change revisions (`w:rPrChange`) with DescribeFormatChange helper
- [x] Table cell revisions (`w:cellIns`, `w:cellDel`, `w:cellMerge`)
- [x] Author coloring (AuthorColors dictionary - CSS already implemented)
- [x] Footnote/endnote rendering (RenderFootnotesAndEndnotes setting)
- [x] Footnote/endnote CSS (section.footnotes, a.footnote-ref, etc.)
- [x] Add unit tests (HC007-HC008)

#### Phase 4: Polish
- [ ] Edge cases (nested revisions, split revisions)
- [ ] Revisions in footnotes/endnotes (requires footnote content transformation)
- [ ] Revisions in headers/footers
- [ ] Header/footer rendering
- [ ] Comprehensive testing
- [ ] Documentation updates

## API Usage Example

```csharp
var settings = new WmlToHtmlConverterSettings
{
    PageTitle = "Document with Tracked Changes",
    RenderTrackedChanges = true,
    IncludeRevisionMetadata = true,
    ShowDeletedContent = true,
    RenderMoveOperations = true,
    RenderFootnotesAndEndnotes = true,
    AuthorColors = new Dictionary<string, string>
    {
        { "John Doe", "#0066cc" },
        { "Jane Smith", "#cc6600" }
    }
};

var html = WmlToHtmlConverter.ConvertToHtml(wmlDoc, settings);
```

## Conclusion

This design provides a comprehensive approach to rendering tracked changes in HTML while:
- Maintaining backward compatibility (default behavior unchanged)
- Supporting all major revision types
- Providing flexible customization options
- Generating semantic HTML (`<ins>`, `<del>`)
- Including machine-readable metadata for tooling
- Producing accessible and visually clear output
