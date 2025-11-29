# Comment Rendering in HTML Converter

This document describes the architecture for rendering Word document comments in HTML output via `WmlToHtmlConverter`.

**Source File:** `Docxodus/WmlToHtmlConverter.cs`

## Overview

Word documents store comments as annotations linked to text ranges. The converter can render these comments in HTML with three different modes:

1. **EndnoteStyle** (default): Comments appear at the end of the document with bidirectional anchor links, similar to footnotes
2. **Inline**: Comments are embedded as `title` attributes and `data-*` attributes for tooltip display
3. **Margin**: Comments are positioned in a side margin using CSS (requires additional styling)

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

```html
<span class="comment-highlight comment-margin-anchor" data-comment-id="0">
  <span>This text has a comment</span>
</span>
<aside class="comment-margin" data-for="0">
  <span class="comment-author">John Doe</span>
  <p>This is the comment text.</p>
</aside>
```

## Generated CSS

When comments are enabled, the following CSS is generated:

```css
/* Comment CSS */
span.comment-highlight {
  background-color: #fff8dc;
  border-bottom: 1px dotted #daa520;
}
sup.comment-marker a {
  color: #0066cc;
  text-decoration: none;
  font-size: 0.75em;
}
aside.comments-section {
  margin-top: 2em;
  padding-top: 1em;
  border-top: 1px solid #ccc;
}
.comment-item {
  margin-bottom: 1em;
  padding: 0.5em;
  background-color: #f9f9f9;
  border-left: 3px solid #daa520;
}
.comment-author {
  font-weight: bold;
  margin-right: 0.5em;
}
.comment-date {
  color: #666;
  font-size: 0.9em;
}
.comment-content {
  margin-top: 0.25em;
}
.comment-backref {
  float: right;
  text-decoration: none;
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
4. **Overlapping comments**: When multiple comments cover the same text, only the first comment's metadata is shown on the highlight span
5. **Margin mode**: Requires custom CSS for proper positioning; no default layout provided

## Future Enhancements

- Support for comment reply threads
- Per-author highlight colors (similar to revision author colors)
- Resolved/active comment state indication
- Improved margin mode with automatic positioning
