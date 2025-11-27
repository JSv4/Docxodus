# DOCX to HTML Converter

This document describes the architecture and capabilities of the `WmlToHtmlConverter` module, which converts Word documents (`.docx`) to XHTML.

**Source File:** `OpenXmlPowerTools/WmlToHtmlConverter.cs`

## Overview

The converter transforms Open XML WordprocessingML documents into well-formed XHTML with CSS styling. It processes the document through a multi-stage pipeline that resolves formatting, calculates layout metrics, and transforms elements to their HTML equivalents.

## Entry Points

```csharp
// From a WmlDocument (in-memory byte array wrapper)
XElement html = WmlToHtmlConverter.ConvertToHtml(wmlDoc, settings);

// From an open WordprocessingDocument
XElement html = WmlToHtmlConverter.ConvertToHtml(wordDoc, settings);

// Via extension method on WmlDocument
XElement html = wmlDoc.ConvertToHtml(settings);
```

All methods return an `XElement` representing the complete HTML document (the `<html>` element with `<head>` and `<body>`).

## Configuration

### WmlToHtmlConverterSettings

| Property | Type | Default | Description |
|----------|------|---------|-------------|
| `PageTitle` | `string` | `""` | Content for the `<title>` element |
| `CssClassPrefix` | `string` | `"pt-"` | Prefix for generated CSS class names |
| `FabricateCssClasses` | `bool` | `true` | If true, generates CSS classes; if false, uses inline styles |
| `GeneralCss` | `string` | `"span { white-space: pre-wrap; }"` | Base CSS included in output |
| `AdditionalCss` | `string` | `""` | Custom CSS appended to the style block |
| `RestrictToSupportedLanguages` | `bool` | `false` | Limits language processing |
| `RestrictToSupportedNumberingFormats` | `bool` | `false` | Limits list numbering formats |
| `ListItemImplementations` | `Dictionary<string, Func<...>>` | Default implementations | Custom list item text generators |
| `ImageHandler` | `Func<ImageInfo, XElement>` | `null` | **Required for images** - callback to process embedded images |

### ImageInfo Structure

When `ImageHandler` is invoked, it receives an `ImageInfo` object:

```csharp
public class ImageInfo
{
    public SKBitmap? Bitmap;           // Decoded image (SkiaSharp)
    public byte[]? ImageBytes;         // Raw image bytes
    public XAttribute? ImgStyleAttribute; // Pre-computed style attribute with dimensions
    public string? ContentType;        // MIME type (e.g., "image/png")
    public XElement? DrawingElement;   // Original Word XML element
    public string? AltText;            // Description from document

    public const int EmusPerInch = 914400;
    public const int EmusPerCm = 360000;

    public void SaveImage(string filePath, SKEncodedImageFormat format, int quality = 100);
}
```

## Processing Pipeline

The conversion happens in the following stages:

### 1. Preprocessing

```
WordprocessingDocument
    │
    ├─► RevisionAccepter.AcceptRevisions()
    │       Accepts all tracked changes (insertions/deletions)
    │
    ├─► MarkupSimplifier.SimplifyMarkup()
    │       Removes: comments, content controls, footnotes/endnotes,
    │                proofing marks, permissions, RSID info, smart tags,
    │                soft hyphens, GoBack bookmarks
    │
    ├─► FormattingAssembler.AssembleFormatting()
    │       Resolves style inheritance, flattens formatting to explicit properties
    │
    ├─► InsertAppropriateNonbreakingSpaces()
    │       Ensures empty paragraphs render (adds space to empty <p>)
    │
    ├─► CalculateSpanWidthForTabs()
    │       Computes tab widths using font metrics (SkiaSharp)
    │
    ├─► ReverseTableBordersForRtlTables()
    │       Swaps left/right borders for RTL tables
    │
    ├─► AdjustTableBorders()
    │       Resolves conflicting cell borders by priority
    │
    ├─► FieldRetriever.AnnotateWithFieldInfo()
    │       Marks field code regions for hyperlink processing
    │
    └─► AnnotateForSections()
            Marks section boundaries for grouping
```

### 2. Transformation

`ConvertToHtmlTransform()` recursively walks the document tree, converting each element:

```
w:document  ──►  <html>
                   <head>
                     <meta charset="UTF-8"/>
                     <title>...</title>
                     <meta name="Generator" content="PowerTools for Open XML"/>
                   </head>
                   <body>...</body>
                 </html>

w:body      ──►  <body> with section <div> wrappers

w:p         ──►  <p> or <h1>-<h6> (based on outlineLvl in style)

w:r         ──►  <span> with CSS styling

w:t         ──►  text node

w:tbl       ──►  <table> wrapped in alignment <div>

w:tr        ──►  <tr>

w:tc        ──►  <td> with colspan/rowspan

w:hyperlink ──►  <a href="...">

w:drawing   ──►  Result of ImageHandler callback
w:pict

w:tab       ──►  <span> with calculated width

w:br/w:cr   ──►  <br> with directional marks

w:sym       ──►  <span>&#NNN;</span> (symbol as entity)

w:bookmarkStart ──► <a id="..."></a>

w:noBreakHyphen ──► "-" (text)
```

### 3. CSS Generation

`ReifyStylesAndClasses()` processes style annotations attached to elements:

**With `FabricateCssClasses = true` (default):**
- Groups elements with identical styles
- Generates CSS class names (e.g., `pt-000001`, or `pt-StyleName` if available)
- Adds `class` attribute to elements
- Creates `<style>` block in `<head>`

**With `FabricateCssClasses = false`:**
- Converts style dictionaries to inline `style` attributes

## Element Mapping Details

### Paragraphs

Paragraphs (`w:p`) are converted based on their style's outline level:

| Outline Level | HTML Element |
|---------------|--------------|
| 0 | `<h1>` |
| 1 | `<h2>` |
| 2 | `<h3>` |
| 3 | `<h4>` |
| 4 | `<h5>` |
| 5 | `<h6>` |
| None/other | `<p>` |

Style separators (`w:specVanish`) cause following paragraphs to be rendered as `<span>` within the heading.

### Runs (Text Formatting)

Run properties (`w:rPr`) map to CSS:

| Word Property | CSS Property | Notes |
|--------------|--------------|-------|
| `w:b` | `font-weight: bold` | |
| `w:i` | `font-style: italic` | |
| `w:u` | `text-decoration: underline` | Unless `val="none"` |
| `w:strike`, `w:dstrike` | `text-decoration: line-through` | |
| `w:caps` | `text-transform: uppercase` | |
| `w:smallCaps` | `font-variant: small-caps` | |
| `w:vanish` | `display: none` | Unless `w:specVanish` |
| `w:color` | `color: #RRGGBB` | `auto` → `black` |
| `w:highlight` | `background: <color>` | Named colors |
| `w:shd` | `background: #RRGGBB` | With pattern blending |
| `w:sz` | `font-size: Npt` | Value in half-points |
| `w:spacing` | `letter-spacing: Npt` | |
| `w:position` | `position: relative; top: Npt` | Superscript/subscript offset |
| `w:vertAlign` | `<sup>` or `<sub>` wrapper | `superscript`/`subscript` |
| `w:bdr` | `border: solid windowtext 1pt` | |

### Tables

Tables are wrapped in a `<div>` for alignment control:

```html
<div dir="ltr" align="left">
  <table dir="ltr" style="border-collapse: collapse; border: none; ...">
    <tr style="height: 0.25in">
      <td colspan="2" rowspan="1" style="border: solid black 1pt; ...">
        <!-- cell content -->
      </td>
    </tr>
  </table>
</div>
```

**Border Conflict Resolution:**

When adjacent cells have different borders, priority is determined by:
1. Border style weight (single < thick < double < dotted, etc.)
2. Border size (`w:sz` attribute)
3. Border style type priority
4. Color value (lower hex value wins)

### Hyperlinks

**External links** (`w:hyperlink` with `r:id`):
```html
<a href="https://example.com">link text</a>
```

**Bookmark links** (`w:hyperlink` with `w:anchor`):
```html
<a href="#bookmarkName" style="text-decoration: none">link text</a>
```

**HYPERLINK fields** (detected via `FieldRetriever`):
```html
<a href="url-from-field">field content</a>
```

### Tabs

Tab characters are converted to `<span>` elements with computed widths:

```html
<span style="margin: 0 0 0 0.75in; padding: 0 0 0 0">&#x00a0;</span>
```

**Tab types supported:**
- `left`/`start` - Content starts at tab position
- `right`/`end` - Content ends at tab position
- `center` - Content centered at tab position
- `decimal` - Decimal point aligned at tab position

**Tab leaders:**
- `dot` (.) `hyphen` (-) `underscore` (_)
- Rendered as repeated characters within the span

### Images

Images require a custom `ImageHandler` callback. Supported formats:
- PNG (`image/png`)
- GIF (`image/gif`)
- TIFF (`image/tiff`)
- JPEG (`image/jpeg`)

**Not supported:**
- WMF (`image/x-wmf`) - Excluded due to memory issues
- EMF - Not processed
- SVG - Not processed

**Example ImageHandler:**

```csharp
settings.ImageHandler = imageInfo =>
{
    // Save image to file
    var filename = $"image_{Guid.NewGuid()}.png";
    imageInfo.SaveImage(filename, SKEncodedImageFormat.Png);

    return new XElement(Xhtml.img,
        new XAttribute("src", filename),
        new XAttribute("alt", imageInfo.AltText ?? ""),
        imageInfo.ImgStyleAttribute);
};
```

### Bidirectional Text (RTL)

The converter handles right-to-left text:

1. **Paragraph-level:** `w:bidi` element triggers `dir="rtl"` and flips margins/alignment
2. **Run-level:** `w:rtl` element adds directional marks (RLM `&#x200f;`)
3. **Table-level:** `w:bidiVisual` reverses left/right borders and alignment

Directional marks ensure proper text rendering:
- LRM (`&#x200e;`) - Left-to-right mark
- RLM (`&#x200f;`) - Right-to-left mark

## CSS Property Mapping

### Paragraph Spacing

| Word Property | CSS Property |
|--------------|--------------|
| `w:spacing/@w:before` | `margin-top` |
| `w:spacing/@w:after` | `margin-bottom` |
| `w:spacing/@w:line` (auto) | `line-height: N%` (where N = line/240*100) |
| `w:spacing/@w:line` (exact) | `line-height: Npt` |

### Indentation

| Word Property | CSS Property |
|--------------|--------------|
| `w:ind/@w:left` | `margin-left` (or `margin-right` for RTL) |
| `w:ind/@w:right` | `margin-right` (or `margin-left` for RTL) |
| `w:ind/@w:firstLine` | `text-indent` |
| `w:ind/@w:hanging` | `text-indent` (negative) |

### Justification

| `w:jc/@w:val` | CSS `text-align` |
|---------------|------------------|
| `left` | `left` (or `right` for RTL) |
| `right` | `right` (or `left` for RTL) |
| `center` | `center` |
| `both` | `justify` |

### Borders

Border styles map to CSS with approximations:

| Word Border | CSS Border |
|-------------|------------|
| `single` | `solid` |
| `dotted` | `dotted` |
| `dashed`, `dashSmallGap`, `dotDash`, `dotDotDash` | `dashed` |
| `double`, `triple` | `double` |
| `thinThick*`, `thickThin*` | `double` |
| `wave`, `doubleWave` | `solid` / `double` |
| `threeDEmboss` | `ridge` |
| `threeDEngrave` | `groove` |
| `outset` | `outset` |
| `inset` | `inset` |

### Colors

Named colors are mapped:

| Word Color | CSS Color |
|------------|-----------|
| `black` | `black` |
| `blue` | `blue` |
| `cyan` | `aqua` |
| `green` | `green` |
| `magenta` | `fuchsia` |
| `red` | `red` |
| `yellow` | `yellow` |
| `white` | `white` |
| `darkBlue` | `#00008B` |
| `darkCyan` | `#008B8B` |
| `darkGreen` | `#006400` |
| `darkMagenta` | `#800080` |
| `darkRed` | `#8B0000` |
| `darkYellow` | `#808000` |
| `darkGray` | `#A9A9A9` |
| `lightGray` | `#D3D3D3` |
| `auto` | `black` (foreground) or `white` (background) |

Hex colors (e.g., `FF0000`) are output as `#FF0000`.

### Font Families

Common fonts include CSS fallbacks:

```css
font-family: 'Arial', 'sans-serif'
font-family: 'Times New Roman', 'serif'
font-family: 'Courier New'
```

## Limitations

### Content Removed During Preprocessing

The following are **discarded** and will not appear in output:

- Comments (`w:comment`)
- Footnotes and endnotes (`w:footnote`, `w:endnote`)
- Content controls (`w:sdt`) - contents preserved, control removed
- Proofing marks (`w:proofErr`)
- Revision tracking info (RSID attributes)
- Smart tags
- Soft hyphens
- Permissions
- GoBack bookmarks

### Not Implemented

**Paragraph properties:**
- `w:framePr` - Frames/text boxes positioning
- `w:keepLines`, `w:keepNext` - Pagination control
- `w:pageBreakBefore` - Page breaks
- `w:widowControl` - Widow/orphan control
- `w:textDirection` - Vertical text
- `w:wordWrap` - Word wrapping control
- `w:mirrorIndents` - Mirrored margins
- `w:suppressOverlap` - Overlap control
- `w:snapToGrid` - Grid snapping

**Run properties:**
- `w:em` - East Asian emphasis marks
- `w:emboss`, `w:imprint` - 3D text effects
- `w:outline`, `w:shadow` - Text effects
- `w:kern` - Kerning
- `w:fitText` - Text scaling
- `w:w` - Character width scaling

**Document features:**
- Math equations (OMML)
- Charts
- Diagrams (SmartArt)
- Complex text boxes
- Headers and footers
- Page numbers
- Table of contents (rendered as static text)
- Cross-references (rendered as static text)
- Multi-column layouts
- Page breaks / section breaks

### Field Code Limitations

Only **HYPERLINK** fields are converted to `<a>` elements. All other field types are rendered as their display text only:
- TOC, INDEX, XE
- REF, PAGEREF, NOTEREF
- DATE, TIME
- AUTHOR, TITLE
- SEQ, LISTNUM
- etc.

### Image Limitations

- **WMF files excluded** - Known to cause memory issues
- **EMF not supported** - Not in allowed content types
- **SVG not supported** - Not in allowed content types
- **Charts not rendered** - Appear as empty space
- **Requires ImageHandler** - Returns `null` without callback

### Font Metrics

Tab width calculation depends on SkiaSharp's font metrics:
- Unknown fonts return zero width (tabs may render incorrectly)
- System font availability affects accuracy
- Azure/server environments may lack fonts

### Browser Compatibility

- Vertical text alignment (`w:textAlignment`) noted as not working in major browsers
- Some border styles are approximations
- Complex RTL layouts may vary by browser

## Output Structure

```html
<!DOCTYPE html>
<html>
  <head>
    <meta charset="UTF-8"/>
    <title>Document Title</title>
    <meta name="Generator" content="PowerTools for Open XML"/>
    <style>
      span { white-space: pre-wrap; }
      p.pt-Normal { margin-top: 0; margin-bottom: 10pt; ... }
      span.pt-DefaultParagraphFont { font-family: 'Calibri'; ... }
      /* ... generated classes ... */
    </style>
  </head>
  <body>
    <div>
      <!-- Section content -->
      <h1 class="pt-Heading1" dir="ltr">...</h1>
      <p class="pt-Normal" dir="ltr">
        <span class="pt-000001">Text content</span>
      </p>
      <div dir="ltr" align="left">
        <table>...</table>
      </div>
    </div>
  </body>
</html>
```

## Usage Example

```csharp
using OpenXmlPowerTools;
using System.Xml.Linq;

// Load document
var wmlDoc = new WmlDocument("document.docx");

// Configure settings
var settings = new WmlToHtmlConverterSettings
{
    PageTitle = "My Document",
    FabricateCssClasses = true,
    AdditionalCss = "body { max-width: 800px; margin: auto; }",
    ImageHandler = imageInfo =>
    {
        // Convert to base64 data URI
        var base64 = Convert.ToBase64String(imageInfo.ImageBytes);
        var dataUri = $"data:{imageInfo.ContentType};base64,{base64}";

        return new XElement(Xhtml.img,
            new XAttribute("src", dataUri),
            new XAttribute("alt", imageInfo.AltText ?? ""),
            imageInfo.ImgStyleAttribute);
    }
};

// Convert
XElement html = WmlToHtmlConverter.ConvertToHtml(wmlDoc, settings);

// Save to file
html.Save("output.html", SaveOptions.DisableFormatting);
```

## Related Files

- `OpenXmlPowerTools/WmlToHtmlConverter.cs` - Main converter implementation
- `OpenXmlPowerTools/HtmlToWmlConverter.cs` - Reverse conversion (HTML to DOCX)
- `OpenXmlPowerTools/RevisionAccepter.cs` - Track changes acceptance
- `OpenXmlPowerTools/MarkupSimplifier.cs` - Document cleanup
- `OpenXmlPowerTools/FormattingAssembler.cs` - Style resolution
- `OpenXmlPowerTools/FieldRetriever.cs` - Field code processing
- `OpenXmlPowerTools/PtOpenXml.cs` - Custom annotation namespace
