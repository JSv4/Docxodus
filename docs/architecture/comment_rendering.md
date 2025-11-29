# Comment Rendering in HTML Converter

This document describes the architecture for rendering Word document comments in HTML output via `WmlToHtmlConverter`.

**Source File:** `Docxodus/WmlToHtmlConverter.cs`

## Overview

Word documents store comments as annotations linked to text ranges. The converter can render these comments in HTML with three different modes:

1. **EndnoteStyle** (default): Comments appear at the end of the document with bidirectional anchor links, similar to footnotes
2. **Inline**: Comments are embedded as `title` attributes and `data-*` attributes for tooltip display
3. **Margin**: Comments are positioned in a side column using CSS flexbox layout

## OOXML Comment Structure

In a `.docx` file, comments are stored across two locations:

### document.xml (Main Document)

Comments are marked in the document body using three elements:

```xml
<w:p>
  <w:commentRangeStart w:id="0"/>
  <w:r>
    <w:t>This text has a comment</w:t>
  </w:r>
  <w:commentRangeEnd w:id="0"/>
  <w:r>
    <w:rPr>
      <w:rStyle w:val="CommentReference"/>
    </w:rPr>
    <w:commentReference w:id="0"/>
  </w:r>
</w:p>
```

| Element | Purpose |
|---------|---------|
| `w:commentRangeStart` | Marks where the commented text begins |
| `w:commentRangeEnd` | Marks where the commented text ends |
| `w:commentReference` | The superscript marker linking to the comment |

### comments.xml (WordprocessingCommentsPart)

The actual comment content is stored separately:

```xml
<w:comments>
  <w:comment w:id="0" w:author="John Doe" w:date="2024-01-15T10:30:00Z" w:initials="JD">
    <w:p>
      <w:r>
        <w:t>This is the comment text.</w:t>
      </w:r>
    </w:p>
  </w:comment>
</w:comments>
```

## Configuration

### WmlToHtmlConverterSettings

| Property | Type | Default | Description |
|----------|------|---------|-------------|
| `RenderComments` | `bool` | `false` | Enable comment rendering |
| `CommentRenderMode` | `CommentRenderMode` | `EndnoteStyle` | How to render comments |
| `CommentCssClassPrefix` | `string` | `"comment-"` | CSS class prefix for comment elements |
| `IncludeCommentMetadata` | `bool` | `true` | Include author/date in output |

### CommentRenderMode Enum

```csharp
public enum CommentRenderMode
{
    EndnoteStyle,  // Comments at end with bidirectional links
    Inline,        // Data attributes on highlighted text
    Margin         // CSS-positioned margin comments
}
```

## Architecture

### Helper Classes

#### CommentInfo

Stores parsed comment data:

```csharp
public class CommentInfo
{
    public int Id { get; set; }
    public string Author { get; set; }
    public string Date { get; set; }
    public string Initials { get; set; }
    public List<XElement> ContentParagraphs { get; set; }
}
```

#### CommentTracker

Tracks comment state during the transformation:

```csharp
internal class CommentTracker
{
    // All comments keyed by ID
    public Dictionary<int, CommentInfo> Comments { get; }

    // Currently open comment ranges (for nested/overlapping comments)
    public HashSet<int> OpenRanges { get; }

    // Comments that were referenced (for rendering section)
    public List<int> ReferencedCommentIds { get; }
}
```

## Processing Pipeline

### 1. Conditional Preservation

When `RenderComments` is enabled, the preprocessing stage preserves comment markup:

```csharp
SimplifyMarkupSettings simplifyMarkupSettings = new SimplifyMarkupSettings
{
    RemoveComments = !htmlConverterSettings.RenderComments,  // Keep if rendering
    // ... other settings
};
```

### 2. Comment Loading

Before transformation, comments are loaded from the `WordprocessingCommentsPart`:

```csharp
private static CommentTracker LoadComments(WordprocessingDocument wordDoc)
{
    var tracker = new CommentTracker();
    var commentsPart = wordDoc.MainDocumentPart?.WordprocessingCommentsPart;

    if (commentsPart != null)
    {
        var commentsDoc = commentsPart.GetXDocument();
        foreach (var commentElement in commentsDoc.Descendants(W.comment))
        {
            var info = new CommentInfo
            {
                Id = (int)commentElement.Attribute(W.id),
                Author = (string)commentElement.Attribute(W.author) ?? "",
                Date = (string)commentElement.Attribute(W.date) ?? "",
                Initials = (string)commentElement.Attribute(W.initials) ?? "",
                ContentParagraphs = commentElement.Elements(W.p).ToList()
            };
            tracker.Comments[info.Id] = info;
        }
    }

    return tracker;
}
```

### 3. Annotation Passing

The `CommentTracker` is passed through the transformation pipeline via an XElement annotation on the root element:

```csharp
// Before transform
mainDocPartXDoc.Root.AddAnnotation(commentTracker);

// During transform - retrieve from ancestors
var tracker = element.Ancestors().First().Annotation<CommentTracker>();
```

### 4. Range Processing

Comment range markers are processed during transformation:

```csharp
// In ConvertToHtmlTransform switch statement
case "commentRangeStart":
    return ProcessCommentRangeStart(element, settings);

case "commentRangeEnd":
    return ProcessCommentRangeEnd(element, settings);

case "commentReference":
    return ProcessCommentReference(element, settings);
```

#### ProcessCommentRangeStart

Opens a comment range by adding the ID to `OpenRanges`:

```csharp
private static object ProcessCommentRangeStart(XElement element, WmlToHtmlConverterSettings settings)
{
    var tracker = element.Ancestors().First().Annotation<CommentTracker>();
    var id = (int?)element.Attribute(W.id);

    if (tracker != null && id.HasValue)
        tracker.OpenRanges.Add(id.Value);

    return null;  // No HTML output
}
```

#### ProcessCommentRangeEnd

Closes a comment range:

```csharp
private static object ProcessCommentRangeEnd(XElement element, WmlToHtmlConverterSettings settings)
{
    var tracker = element.Ancestors().First().Annotation<CommentTracker>();
    var id = (int?)element.Attribute(W.id);

    if (tracker != null && id.HasValue)
        tracker.OpenRanges.Remove(id.Value);

    return null;  // No HTML output
}
```

#### ProcessCommentReference

Creates the superscript marker linking to the comment:

```csharp
private static object ProcessCommentReference(XElement element, WmlToHtmlConverterSettings settings)
{
    var tracker = element.Ancestors().First().Annotation<CommentTracker>();
    var id = (int?)element.Attribute(W.id);
    var prefix = settings.CommentCssClassPrefix ?? "comment-";

    if (tracker != null && id.HasValue && tracker.Comments.ContainsKey(id.Value))
    {
        tracker.ReferencedCommentIds.Add(id.Value);

        // EndnoteStyle: Create anchor link
        if (settings.CommentRenderMode == CommentRenderMode.EndnoteStyle)
        {
            return new XElement(Xhtml.sup,
                new XAttribute("class", prefix + "marker"),
                new XAttribute("id", prefix + "ref-" + id.Value),
                new XElement(Xhtml.a,
                    new XAttribute("href", "#" + prefix.TrimEnd('-') + "-" + id.Value),
                    "[" + (tracker.ReferencedCommentIds.Count) + "]"
                )
            );
        }
    }

    return null;
}
```

### 5. Text Highlighting

When text runs are within an open comment range, they are wrapped in highlight spans:

```csharp
// In ConvertRun, after creating the span
if (tracker?.OpenRanges.Count > 0)
{
    var prefix = settings.CommentCssClassPrefix ?? "comment-";
    var commentIds = tracker.OpenRanges.ToList();

    // Wrap in highlight span
    var highlightSpan = new XElement(Xhtml.span,
        new XAttribute("class", prefix + "highlight"),
        new XAttribute("data-comment-id", string.Join(",", commentIds))
    );

    if (settings.IncludeCommentMetadata)
    {
        var firstComment = tracker.Comments[commentIds[0]];
        highlightSpan.Add(new XAttribute("data-author", firstComment.Author));
    }

    // For Inline mode, add tooltip attributes
    if (settings.CommentRenderMode == CommentRenderMode.Inline)
    {
        var commentText = GetCommentPlainText(tracker.Comments[commentIds[0]]);
        highlightSpan.Add(
            new XAttribute("title", firstComment.Author + ": " + commentText),
            new XAttribute("data-comment", commentText)
        );
    }

    highlightSpan.Add(runSpan);
    return highlightSpan;
}
```

### 6. Comments Section Rendering

For `EndnoteStyle` mode, a comments section is appended to the document body:

```csharp
private static XElement RenderCommentsSection(CommentTracker tracker, WmlToHtmlConverterSettings settings)
{
    var prefix = settings.CommentCssClassPrefix ?? "comment-";
    var sectionName = prefix.TrimEnd('-') + "s";  // "comments" or custom

    var section = new XElement(Xhtml.aside,
        new XAttribute("class", sectionName + "-section"),
        new XElement(Xhtml.h2, "Comments")
    );

    foreach (var id in tracker.ReferencedCommentIds)
    {
        if (tracker.Comments.TryGetValue(id, out var comment))
        {
            section.Add(RenderCommentItem(comment, id, settings));
        }
    }

    return section;
}
```

Each comment item includes:
- Anchor ID for linking
- Author and date metadata
- Comment content (paragraphs)
- Back-reference link to the text

## HTML Output Examples

### EndnoteStyle Mode

```html
<!-- In document body -->
<span class="comment-highlight" data-comment-id="0" data-author="John Doe">
  <span>This text has a comment</span>
</span>
<sup class="comment-marker" id="comment-ref-0">
  <a href="#comment-0">[1]</a>
</sup>

<!-- At end of document -->
<aside class="comments-section">
  <h2>Comments</h2>
  <div class="comment-item" id="comment-0">
    <span class="comment-author">John Doe</span>
    <span class="comment-date">2024-01-15</span>
    <div class="comment-content">
      <p>This is the comment text.</p>
    </div>
    <a class="comment-backref" href="#comment-ref-0">↩</a>
  </div>
</aside>
```

### Inline Mode

```html
<span class="comment-highlight"
      data-comment-id="0"
      data-author="John Doe"
      data-comment="This is the comment text."
      title="John Doe: This is the comment text.">
  <span>This text has a comment</span>
</span>
```

### Margin Mode

In margin mode, the entire document body is wrapped in a flexbox container with the main content on the left and a comment column on the right:

```html
<div class="comment-margin-container">
  <!-- Main content area -->
  <div class="comment-margin-content">
    <p>
      <span class="comment-highlight" data-comment-id="0">
        This text has a comment
      </span>
      <sup><a href="#comment-0" class="comment-marker" id="comment-ref-0">[0]</a></sup>
    </p>
  </div>

  <!-- Margin column with comments -->
  <aside class="comment-margin-column">
    <div class="comment-margin-note" id="comment-0" data-comment-id="0">
      <div class="comment-margin-note-header">
        <span class="comment-margin-author">John Doe</span>
        <span class="comment-margin-date">Jan 15</span>
        <a href="#comment-ref-0" class="comment-margin-backref">↩</a>
      </div>
      <div class="comment-margin-note-body">
        <p>This is the comment text.</p>
      </div>
    </div>
  </aside>
</div>
```

## Generated CSS

When comments are enabled, base CSS is generated for all modes. Additional CSS is generated for margin mode.

### Base CSS (All Modes)

```css
/* Comments CSS */
span.comment-highlight {
  background-color: #fff9c4;
  border-bottom: 2px solid #fbc02d;
}
a.comment-marker {
  color: #1976d2;
  text-decoration: none;
  margin-left: 2px;
}
a.comment-marker:hover {
  text-decoration: underline;
}
```

### EndnoteStyle Mode CSS

```css
aside.comments-section {
  margin-top: 2em;
  padding-top: 1em;
  border-top: 2px solid #ccc;
}
aside.comments-section h2 {
  font-size: 1.2em;
  margin-bottom: 0.5em;
}
ol.comments-list {
  list-style: none;
  padding: 0;
}
li.comment {
  margin-bottom: 1em;
  padding: 0.75em;
  background-color: #f5f5f5;
  border-left: 3px solid #1976d2;
  border-radius: 0 4px 4px 0;
}
div.comment-header {
  display: flex;
  align-items: center;
  gap: 0.5em;
  margin-bottom: 0.5em;
  font-size: 0.85em;
}
span.comment-author {
  font-weight: bold;
  color: #1976d2;
}
span.comment-date {
  color: #666;
}
a.comment-backref {
  margin-left: auto;
  text-decoration: none;
  color: #1976d2;
}
div.comment-body p {
  margin: 0;
}
```

### Margin Mode CSS

When `CommentRenderMode.Margin` is selected, additional flexbox layout CSS is generated:

```css
/* Margin Mode Comments */
div.comment-margin-container {
  display: flex;
  flex-direction: row;
  gap: 1em;
}
div.comment-margin-content {
  flex: 1;
  min-width: 0;
}
aside.comment-margin-column {
  width: 250px;
  flex-shrink: 0;
  position: relative;
}
div.comment-margin-note {
  position: relative;
  margin-bottom: 0.5em;
  padding: 0.5em;
  background-color: #fff9c4;
  border-left: 3px solid #fbc02d;
  border-radius: 0 4px 4px 0;
  font-size: 0.85em;
  box-shadow: 0 1px 3px rgba(0,0,0,0.1);
}
div.comment-margin-note-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 0.25em;
  font-size: 0.9em;
}
span.comment-margin-author {
  font-weight: bold;
  color: #f57f17;
}
span.comment-margin-date {
  color: #666;
  font-size: 0.85em;
}
div.comment-margin-note-body {
  color: #333;
}
div.comment-margin-note-body p {
  margin: 0;
}
a.comment-margin-backref {
  color: #1976d2;
  text-decoration: none;
  font-size: 0.85em;
}
a.comment-margin-backref:hover {
  text-decoration: underline;
}
span.comment-highlight[data-comment-id] {
  cursor: pointer;
}

/* Print styles for margin mode */
@media print {
  div.comment-margin-container {
    display: block;
  }
  aside.comment-margin-column {
    width: auto;
    page-break-inside: avoid;
  }
}
```

## WASM/npm Integration

The npm package exposes comment options via `ConversionOptions`:

```typescript
import { convertDocxToHtml, CommentRenderMode } from 'docxodus';

// Render comments in endnote style
const html = await convertDocxToHtml(docxFile, {
  commentRenderMode: CommentRenderMode.EndnoteStyle,
  commentCssClassPrefix: 'note-',
});

// Render comments as inline tooltips
const htmlInline = await convertDocxToHtml(docxFile, {
  commentRenderMode: CommentRenderMode.Inline,
});

// Disable comment rendering (default)
const htmlNoComments = await convertDocxToHtml(docxFile, {
  commentRenderMode: CommentRenderMode.Disabled,
});
```

The `CommentRenderMode` enum values:
- `Disabled` (-1): Do not render comments (default)
- `EndnoteStyle` (0): Comments at end with bidirectional links
- `Inline` (1): Comments as tooltips with data attributes
- `Margin` (2): CSS-positioned margin comments

## Limitations

1. **Reply threads**: Word supports threaded comment replies, but these are flattened in the current implementation
2. **Resolved comments**: The resolved/done state of comments is not currently rendered
3. **Comment highlighting colors**: Word allows different highlight colors per comment; currently all use the same CSS
4. **Overlapping comments**: When multiple comments cover the same text, they are nested; only the innermost comment's metadata is most visible
5. **Margin mode positioning**: Comments in margin mode are rendered in document order, not positioned adjacent to their anchor text (would require JavaScript for dynamic positioning)

## Future Enhancements

- Support for comment reply threads
- Per-author highlight colors (similar to revision author colors)
- Resolved/active comment state indication
- JavaScript-based dynamic positioning for margin mode comments
