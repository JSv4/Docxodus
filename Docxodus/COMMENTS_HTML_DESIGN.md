# Comments HTML Rendering - Design Document

## Executive Summary

This document outlines a design for extending `WmlToHtmlConverter` to support rendering Word document comments in the HTML output.

## Current Behavior

The current `WmlToHtmlConverter.ConvertToHtml()` method (line 289) sets:

```csharp
SimplifyMarkupSettings simplifyMarkupSettings = new SimplifyMarkupSettings
{
    RemoveComments = true,
    // ...
};
```

This **removes all comments** before conversion, meaning:
- Comment range markers (`w:commentRangeStart`, `w:commentRangeEnd`) are stripped
- Comment references (`w:commentReference`) are removed
- The comments part (`WordprocessingCommentsPart`) is deleted entirely

## Comment Elements in OOXML

### In the Main Document (document.xml)

| Element | Description | Attributes |
|---------|-------------|------------|
| `w:commentRangeStart` | Start marker for commented text | `w:id` - links to comment |
| `w:commentRangeEnd` | End marker for commented text | `w:id` - links to comment |
| `w:commentReference` | Reference marker (usually superscript) | `w:id` - links to comment |

The typical structure in a paragraph:

```xml
<w:p>
  <w:commentRangeStart w:id="0"/>
  <w:r>
    <w:t>commented text</w:t>
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

### In the Comments Part (comments.xml)

```xml
<w:comments>
  <w:comment w:id="0" w:author="John Doe" w:date="2024-01-15T10:30:00Z" w:initials="JD">
    <w:p>
      <w:pPr>
        <w:pStyle w:val="CommentText"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:rStyle w:val="CommentReference"/>
        </w:rPr>
        <w:annotationRef/>
      </w:r>
      <w:r>
        <w:t>This is the comment text.</w:t>
      </w:r>
    </w:p>
  </w:comment>
</w:comments>
```

### Comment Attributes

| Attribute | Description |
|-----------|-------------|
| `w:id` | Unique comment ID (integer) |
| `w:author` | Author name |
| `w:date` | ISO 8601 timestamp |
| `w:initials` | Author initials |

### Extended Comments (commentsExtended.xml)

Modern Word documents may also have a `WordprocessingCommentsExPart` with:
- `w15:commentEx` elements for threaded replies
- `w15:paraId` for paragraph-level linking
- `w15:done` for resolved comments

## Proposed Design

### 1. Settings Extension

Add new properties to `WmlToHtmlConverterSettings`:

```csharp
public class WmlToHtmlConverterSettings
{
    // ... existing properties ...

    /// <summary>
    /// If true, render comments in HTML output.
    /// If false (default), comments are stripped from the output.
    /// </summary>
    public bool RenderComments;

    /// <summary>
    /// CSS class prefix for comment elements (default: "comment-")
    /// </summary>
    public string CommentCssClassPrefix;

    /// <summary>
    /// How to render comments in the HTML output.
    /// </summary>
    public CommentRenderMode CommentRenderMode;

    /// <summary>
    /// If true, include comment metadata (author, date) as data attributes
    /// </summary>
    public bool IncludeCommentMetadata;
}

public enum CommentRenderMode
{
    /// <summary>
    /// Comments are rendered as a sidebar/aside at the end of the document
    /// with links to/from commented text (default).
    /// </summary>
    EndnoteStyle,

    /// <summary>
    /// Comments are rendered inline as tooltips/popups on the highlighted text.
    /// Uses data attributes and CSS/JS for display.
    /// </summary>
    Inline,

    /// <summary>
    /// Comments are rendered in a margin area using CSS positioning.
    /// Best for print-style layouts.
    /// </summary>
    Margin
}
```

### 2. HTML Output Format

#### EndnoteStyle Mode (Default)

Highlighted text in the document:
```html
<span class="comment-highlight" id="comment-ref-1" data-comment-id="1">
  commented text
  <a href="#comment-1" class="comment-marker" title="Comment by John Doe">[1]</a>
</span>
```

Comments section at the end:
```html
<aside class="comments-section">
  <h2>Comments</h2>
  <ol class="comments-list">
    <li id="comment-1" class="comment" data-author="John Doe" data-date="2024-01-15T10:30:00Z">
      <div class="comment-header">
        <span class="comment-author">John Doe</span>
        <span class="comment-date">Jan 15, 2024</span>
        <a href="#comment-ref-1" class="comment-backref">↩</a>
      </div>
      <div class="comment-body">
        <p>This is the comment text.</p>
      </div>
    </li>
  </ol>
</aside>
```

#### Inline Mode

Highlighted text with embedded comment:
```html
<span class="comment-highlight"
      data-comment-id="1"
      data-author="John Doe"
      data-date="2024-01-15T10:30:00Z"
      data-comment="This is the comment text."
      title="John Doe: This is the comment text.">
  commented text
</span>
```

#### Margin Mode

```html
<span class="comment-highlight comment-anchor" data-comment-id="1">
  commented text
</span>
<!-- Comment rendered via CSS in margin -->
<aside class="comment-margin" data-for="1">
  <div class="comment-bubble">
    <span class="comment-author">JD</span>
    <p>This is the comment text.</p>
  </div>
</aside>
```

### 3. Default CSS

```css
/* Comment Highlights */
.comment-highlight {
  background-color: #fff9c4; /* light yellow */
  border-bottom: 2px solid #fbc02d;
}

.comment-marker {
  color: #1976d2;
  font-size: 0.75em;
  vertical-align: super;
  text-decoration: none;
  margin-left: 2px;
}

.comment-marker:hover {
  text-decoration: underline;
}

/* Comments Section (EndnoteStyle) */
.comments-section {
  margin-top: 2em;
  padding-top: 1em;
  border-top: 2px solid #ccc;
}

.comments-section h2 {
  font-size: 1.2em;
  margin-bottom: 0.5em;
}

.comments-list {
  list-style: none;
  padding: 0;
}

.comment {
  margin-bottom: 1em;
  padding: 0.75em;
  background-color: #f5f5f5;
  border-left: 3px solid #1976d2;
  border-radius: 0 4px 4px 0;
}

.comment-header {
  display: flex;
  align-items: center;
  gap: 0.5em;
  margin-bottom: 0.5em;
  font-size: 0.85em;
}

.comment-author {
  font-weight: bold;
  color: #1976d2;
}

.comment-date {
  color: #666;
}

.comment-backref {
  margin-left: auto;
  text-decoration: none;
  color: #1976d2;
}

.comment-body p {
  margin: 0;
}

/* Inline Mode (tooltip) */
.comment-highlight[data-comment] {
  cursor: help;
}

/* Margin Mode */
.comment-margin {
  position: absolute;
  right: -250px;
  width: 220px;
  font-size: 0.85em;
}

.comment-bubble {
  background: #fff9c4;
  border: 1px solid #fbc02d;
  border-radius: 4px;
  padding: 8px;
  box-shadow: 0 2px 4px rgba(0,0,0,0.1);
}
```

### 4. Implementation Changes

#### 4.1. Entry Point Modification

```csharp
public static XElement ConvertToHtml(WordprocessingDocument wordDoc,
    WmlToHtmlConverterSettings htmlConverterSettings)
{
    // ... existing revision handling ...

    SimplifyMarkupSettings simplifyMarkupSettings = new SimplifyMarkupSettings
    {
        RemoveComments = !htmlConverterSettings.RenderComments, // Changed!
        // ... rest of settings ...
    };
    MarkupSimplifier.SimplifyMarkup(wordDoc, simplifyMarkupSettings);

    // ... rest of existing code ...
}
```

#### 4.2. Track Comment Ranges

Create a helper class to track comment ranges during conversion:

```csharp
private class CommentTracker
{
    public Dictionary<int, CommentInfo> Comments { get; } = new();
    public HashSet<int> OpenRanges { get; } = new();
}

private class CommentInfo
{
    public int Id { get; set; }
    public string Author { get; set; }
    public string Date { get; set; }
    public string Initials { get; set; }
    public List<XElement> ContentParagraphs { get; set; } = new();
}
```

#### 4.3. Load Comments Part

```csharp
private static Dictionary<int, CommentInfo> LoadComments(WordprocessingDocument wordDoc)
{
    var comments = new Dictionary<int, CommentInfo>();
    var commentsPart = wordDoc.MainDocumentPart.WordprocessingCommentsPart;
    if (commentsPart == null)
        return comments;

    var commentsXDoc = commentsPart.GetXDocument();
    foreach (var comment in commentsXDoc.Root?.Elements(W.comment) ?? Enumerable.Empty<XElement>())
    {
        var id = (int?)comment.Attribute(W.id);
        if (id == null) continue;

        comments[id.Value] = new CommentInfo
        {
            Id = id.Value,
            Author = (string)comment.Attribute(W.author),
            Date = (string)comment.Attribute(W.date),
            Initials = (string)comment.Attribute(W.initials),
            ContentParagraphs = comment.Elements(W.p).ToList()
        };
    }

    return comments;
}
```

#### 4.4. New Handler Methods

Add handlers in `ConvertToHtmlTransform()`:

```csharp
// Handle w:commentRangeStart
if (element.Name == W.commentRangeStart)
{
    return ProcessCommentRangeStart(wordDoc, settings, element, commentTracker);
}

// Handle w:commentRangeEnd
if (element.Name == W.commentRangeEnd)
{
    return ProcessCommentRangeEnd(wordDoc, settings, element, commentTracker);
}

// Handle w:commentReference
if (element.Name == W.commentReference)
{
    return ProcessCommentReference(wordDoc, settings, element, commentTracker);
}
```

#### 4.5. Comment Range Processing

The challenge with comment ranges is that they can span multiple runs, paragraphs, and even tables. Options:

**Option A: Wrap each run individually** (simplest)
```html
<span class="comment-highlight" data-comment-id="1">text in run 1</span>
<span class="comment-highlight" data-comment-id="1">text in run 2</span>
```

**Option B: Post-process to merge adjacent highlights** (cleaner output)
Use a transform pass after initial conversion to merge adjacent spans with the same comment ID.

**Option C: Track range and wrap block-level** (most complex)
Track open ranges and wrap at paragraph level when possible.

Recommendation: Start with Option A for simplicity, with the infrastructure to support Option B later.

```csharp
private static object ProcessCommentRangeStart(WordprocessingDocument wordDoc,
    WmlToHtmlConverterSettings settings, XElement element, CommentTracker tracker)
{
    if (!settings.RenderComments)
        return null;

    var id = (int?)element.Attribute(W.id);
    if (id != null)
        tracker.OpenRanges.Add(id.Value);

    // Don't emit anything - we'll wrap content in ConvertRun
    return null;
}

private static object ProcessCommentRangeEnd(WordprocessingDocument wordDoc,
    WmlToHtmlConverterSettings settings, XElement element, CommentTracker tracker)
{
    if (!settings.RenderComments)
        return null;

    var id = (int?)element.Attribute(W.id);
    if (id != null)
        tracker.OpenRanges.Remove(id.Value);

    return null;
}
```

#### 4.6. Modify ConvertRun to Apply Comment Highlighting

```csharp
private static object ConvertRun(WordprocessingDocument wordDoc,
    WmlToHtmlConverterSettings settings, XElement run, CommentTracker tracker)
{
    // ... existing run conversion logic ...

    var result = /* existing span element */;

    if (settings.RenderComments && tracker.OpenRanges.Any())
    {
        // Wrap in comment highlight span(s)
        foreach (var commentId in tracker.OpenRanges)
        {
            var wrapper = new XElement(Xhtml.span,
                new XAttribute("class", (settings.CommentCssClassPrefix ?? "comment-") + "highlight"),
                new XAttribute("data-comment-id", commentId.ToString()));

            wrapper.Add(result);
            result = wrapper;
        }
    }

    return result;
}
```

#### 4.7. Comment Reference (Marker) Processing

```csharp
private static object ProcessCommentReference(WordprocessingDocument wordDoc,
    WmlToHtmlConverterSettings settings, XElement element, CommentTracker tracker)
{
    if (!settings.RenderComments)
        return null;

    var id = (int?)element.Attribute(W.id);
    if (id == null)
        return null;

    var comment = tracker.Comments.GetValueOrDefault(id.Value);
    var prefix = settings.CommentCssClassPrefix ?? "comment-";

    var marker = new XElement(Xhtml.a,
        new XAttribute("href", $"#comment-{id}"),
        new XAttribute("class", prefix + "marker"),
        new XAttribute("id", $"comment-ref-{id}"));

    if (comment != null && settings.IncludeCommentMetadata)
    {
        marker.Add(new XAttribute("title", $"Comment by {comment.Author}"));
    }

    marker.Add(new XText($"[{id}]"));

    return marker;
}
```

#### 4.8. Render Comments Section

```csharp
private static XElement RenderCommentsSection(WordprocessingDocument wordDoc,
    WmlToHtmlConverterSettings settings, CommentTracker tracker)
{
    if (!settings.RenderComments || !tracker.Comments.Any())
        return null;

    var prefix = settings.CommentCssClassPrefix ?? "comment-";

    var section = new XElement(Xhtml.aside,
        new XAttribute("class", prefix.TrimEnd('-') + "s-section"),
        new XElement(Xhtml.h2, "Comments"),
        new XElement(Xhtml.ol,
            new XAttribute("class", prefix.TrimEnd('-') + "s-list"),
            tracker.Comments.Values.OrderBy(c => c.Id).Select(c =>
                RenderCommentItem(wordDoc, settings, c, prefix))));

    return section;
}

private static XElement RenderCommentItem(WordprocessingDocument wordDoc,
    WmlToHtmlConverterSettings settings, CommentInfo comment, string prefix)
{
    var li = new XElement(Xhtml.li,
        new XAttribute("id", $"comment-{comment.Id}"),
        new XAttribute("class", prefix.TrimEnd('-')));

    if (settings.IncludeCommentMetadata)
    {
        li.Add(new XAttribute("data-author", comment.Author ?? ""));
        if (comment.Date != null)
            li.Add(new XAttribute("data-date", comment.Date));
    }

    // Header with author, date, and back link
    var header = new XElement(Xhtml.div,
        new XAttribute("class", prefix + "header"));

    if (comment.Author != null)
        header.Add(new XElement(Xhtml.span,
            new XAttribute("class", prefix + "author"),
            comment.Author));

    if (comment.Date != null)
    {
        // Format date nicely
        if (DateTime.TryParse(comment.Date, out var dt))
        {
            header.Add(new XElement(Xhtml.span,
                new XAttribute("class", prefix + "date"),
                dt.ToString("MMM d, yyyy")));
        }
    }

    header.Add(new XElement(Xhtml.a,
        new XAttribute("href", $"#comment-ref-{comment.Id}"),
        new XAttribute("class", prefix + "backref"),
        "↩"));

    li.Add(header);

    // Comment body - convert paragraphs
    var body = new XElement(Xhtml.div,
        new XAttribute("class", prefix + "body"));

    foreach (var para in comment.ContentParagraphs)
    {
        // Skip the annotation reference run
        var textContent = para.Descendants(W.t)
            .Where(t => t.Ancestors(W.r)
                .All(r => r.Elements(W.annotationRef).Count() == 0))
            .Select(t => t.Value)
            .StringConcatenate();

        if (!string.IsNullOrWhiteSpace(textContent))
            body.Add(new XElement(Xhtml.p, textContent));
    }

    li.Add(body);

    return li;
}
```

#### 4.9. Add Comments Section to Body

Modify the body transformation:

```csharp
if (element.Name == W.body)
{
    var bodyContent = new List<object>();

    // ... existing header rendering ...

    // Main content
    bodyContent.Add(CreateSectionDivs(wordDoc, settings, element));

    // ... existing footnotes/endnotes rendering ...

    // Add comments section if enabled
    if (settings.RenderComments)
    {
        var commentsSection = RenderCommentsSection(wordDoc, settings, commentTracker);
        if (commentsSection != null)
            bodyContent.Add(commentsSection);
    }

    // ... existing footer rendering ...

    return new XElement(Xhtml.body, bodyContent);
}
```

### 5. Threading Context

For the CommentTracker to be accessible throughout the transform, we need to:

**Option A: Add to settings object** (not ideal - settings should be immutable)

**Option B: Use an annotation on the root element**

**Option C: Pass as parameter through transform chain** (recommended)

We'll need to modify `ConvertToHtmlTransform` signature and all calling code:

```csharp
private static object ConvertToHtmlTransform(
    WordprocessingDocument wordDoc,
    WmlToHtmlConverterSettings settings,
    XNode node,
    bool suppressTrailingWhiteSpace,
    decimal currentMarginLeft,
    CommentTracker commentTracker)  // New parameter
```

### 6. Edge Cases

1. **Overlapping comment ranges**: Multiple comments on same text - use nested spans
2. **Comments spanning paragraphs**: Track range state across paragraph boundaries
3. **Comments spanning tables**: Similar to paragraph spanning
4. **Comments in headers/footers**: Process comment parts in those parts too
5. **Comments in footnotes/endnotes**: Similar handling
6. **Nested comments**: Rare but possible - nested highlighting
7. **Empty comment ranges**: No text between start/end markers
8. **Comments with no text content**: Just a marker, no highlighted text
9. **Threaded replies**: Extended comments (Phase 2 feature)

### 7. Testing Strategy

Create test cases for:

1. Simple single comment
2. Multiple comments by same author
3. Multiple comments by different authors
4. Comment spanning multiple runs
5. Comment spanning multiple paragraphs
6. Overlapping comments
7. Comments in tables
8. Comments in headers/footers
9. Comments in footnotes
10. Empty comment (no highlighted text)
11. Long multi-paragraph comment content
12. Comment with formatting in content
13. All render modes (EndnoteStyle, Inline, Margin)

### 8. Implementation Phases

#### Phase 1: Core Infrastructure
- [ ] Add settings properties
- [ ] Skip comment removal when `RenderComments = true`
- [ ] Load comments from comments part
- [ ] Track comment ranges during transform
- [ ] Highlight commented text
- [ ] Render comment markers
- [ ] Render comments section (EndnoteStyle)
- [ ] Generate CSS
- [ ] Basic unit tests

#### Phase 2: Enhanced Rendering
- [ ] Inline mode with data attributes
- [ ] Margin mode with CSS positioning
- [ ] Author-specific coloring
- [ ] Merge adjacent comment highlights
- [ ] Handle comments in headers/footers/footnotes

#### Phase 3: Extended Features
- [ ] Threaded comment replies (CommentsExPart)
- [ ] Resolved/done comments
- [ ] Print-friendly comment output
- [ ] Interactive JS behaviors (optional)

## API Usage Example

```csharp
var settings = new WmlToHtmlConverterSettings
{
    PageTitle = "Document with Comments",
    RenderComments = true,
    CommentRenderMode = CommentRenderMode.EndnoteStyle,
    IncludeCommentMetadata = true,
    CommentCssClassPrefix = "comment-"
};

var html = WmlToHtmlConverter.ConvertToHtml(wmlDoc, settings);
```

## Conclusion

This design provides a comprehensive approach to rendering comments in HTML while:
- Maintaining backward compatibility (default behavior unchanged)
- Supporting multiple rendering modes for different use cases
- Preserving comment metadata for tooling
- Generating accessible and visually clear output
- Handling edge cases gracefully
