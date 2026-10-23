# Unsupported Content Placeholders

*Last updated: December 2025*

This document describes the placeholder rendering feature for unsupported content types in the WmlToHtmlConverter.

## Overview

By default, content that cannot be fully converted to HTML (such as WMF images, math equations, and form fields) is silently dropped from the output. This can make it difficult for users to know that content is missing.

The **Unsupported Content Placeholders** feature provides visual indicators for this content, allowing users to see what content types were encountered but not fully rendered.

## Enabling Placeholders

Placeholders are **disabled by default** to preserve backward compatibility. To enable them:

### C# (.NET)

```csharp
var settings = new WmlToHtmlConverterSettings
{
    RenderUnsupportedContentPlaceholders = true,
    UnsupportedContentCssClassPrefix = "unsupported-",  // optional, this is the default
    IncludeUnsupportedContentMetadata = true,           // optional, this is the default
};

var html = WmlToHtmlConverter.ConvertToHtml(wordDoc, settings);
```

### TypeScript/JavaScript (npm)

```typescript
import { convertDocxToHtml } from 'docxodus';

const html = await convertDocxToHtml(docxBytes, {
    renderUnsupportedContentPlaceholders: true,
});
```

## Supported Content Types

| Content Type | XML Element(s) | Placeholder Text | CSS Class |
|--------------|----------------|------------------|-----------|
| WMF Image | `image/x-wmf` | `[WMF IMAGE]` | `unsupported-image` |
| EMF Image | `image/x-emf` | `[EMF IMAGE]` | `unsupported-image` |
| SVG Image | `image/svg+xml` | `[SVG IMAGE]` | `unsupported-image` |
| Math Equation | `m:oMath`, `m:oMathPara` | `[MATH]` | `unsupported-math` |
| Checkbox | `w:ffData` with `w:checkBox` | `[CHECKBOX]` | `unsupported-form` |
| Text Input | `w:ffData` with `w:textInput` | `[TEXT INPUT]` | `unsupported-form` |
| Dropdown | `w:ffData` with `w:ddList` | `[DROPDOWN]` | `unsupported-form` |
| Ruby Annotation | `w:ruby` | `{base text}` | `unsupported-ruby` |
| Inline drawing the converter cannot draw | `w:drawing/wp:inline` | `[SHAPE]`, `[SMARTART]`, `[CHART]`, `[UNSUPPORTED IMAGE]` or `[DRAWING]` | `unsupported-other` |

### Inline drawings

An inline drawing occupies its `wp:extent` whether or not the converter can draw it (issue #1270): the
empty-paragraph pass counts an inline `w:drawing` as paragraph content, so a drawing that rendered to
nothing used to leave its paragraph with no line box. When nothing in the drawing can be drawn — preset
geometry other than `rect`, `roundRect`, `ellipse`, `line` and `straightConnector1`, custom geometry,
SmartArt, a chart type that is not projected, a picture whose image is missing or linked — the converter
emits:

- with placeholders **on**, the placeholder above (`data-content-type="Other"`, `data-element-name` the
  drawing's content element: `wsp`, `relIds`, `chart`, `pic`, …), sized to the extent with
  `box-sizing: border-box` and marked `data-docx-shape="unrendered"`;
- with placeholders **off**, a blank `<span data-docx-shape="unrendered" aria-hidden="true">` holding an
  empty SVG of the extent.

A shape or group that is only partly drawable (a group holding a picture or a triangle beside a
rectangle, a grouped text box's text, a picture fill or an unsupported fill colour) draws what it can and marks its wrapper
`data-docx-shape="partial"`; with placeholders on, the wrapper also carries `data-content-type="Other"`.
A picture is still omitted entirely when the caller passes no image handler; an image handler that
returns null declines that one picture, whose extent is kept as an unrendered box.

The PDF export always converts with placeholders on and inventories every `data-content-type` element,
so an undrawable drawing is reported as an `unsupported_content_placeholder` warning (and in the render
report's `unsupportedContent`) and rejects the export under `unsupportedContent: 'strict'`.

## HTML Output

Each placeholder is rendered as a `<span>` element with:

```html
<span class="unsupported-placeholder unsupported-math"
      data-content-type="MathEquation"
      data-element-name="oMath"
      title="Math equation (Office Math Markup) - not supported in HTML output">
  [MATH]
</span>
```

### Attributes

| Attribute | Description |
|-----------|-------------|
| `class` | Contains `unsupported-placeholder` plus a type-specific class (e.g., `unsupported-math`) |
| `data-content-type` | The `UnsupportedContentType` enum value (e.g., `MathEquation`, `WmfImage`) |
| `data-element-name` | The local name of the XML element that was not converted |
| `title` | Human-readable description shown on hover |

## CSS Styling

When placeholders are enabled, CSS is automatically generated and included in the HTML output:

```css
/* Unsupported Content Placeholders CSS */
.unsupported-placeholder {
    display: inline-block;
    background-color: #fff3cd;
    border: 1px dashed #856404;
    border-radius: 3px;
    padding: 2px 6px;
    font-family: monospace;
    font-size: 0.85em;
    color: #856404;
    cursor: help;
    vertical-align: middle;
}

/* Type-specific colors */
.unsupported-image { background-color: #d4edda; border-color: #28a745; color: #155724; }
.unsupported-math { background-color: #d1ecf1; border-color: #17a2b8; color: #0c5460; }
.unsupported-form { background-color: #e2e3e5; border-color: #6c757d; color: #383d41; }
.unsupported-ruby { background-color: #cce5ff; border-color: #0d6efd; color: #084298; }
.unsupported-object { background-color: #e2d9f3; border-color: #6f42c1; color: #432874; }
```

### Customizing Colors

You can override the default colors by adding your own CSS after the generated styles:

```csharp
var settings = new WmlToHtmlConverterSettings
{
    RenderUnsupportedContentPlaceholders = true,
    AdditionalCss = @"
        .unsupported-placeholder { background: #f0f0f0; border-color: #999; color: #666; }
        .unsupported-math { background: #ffe6e6; border-color: #ff0000; color: #990000; }
    "
};
```

### Custom CSS Prefix

You can change the CSS class prefix:

```csharp
var settings = new WmlToHtmlConverterSettings
{
    RenderUnsupportedContentPlaceholders = true,
    UnsupportedContentCssClassPrefix = "missing-",  // produces .missing-placeholder, .missing-math, etc.
};
```

## TypeScript Enum

For type-safe identification of unsupported content types in TypeScript:

```typescript
import { UnsupportedContentType } from 'docxodus';

// Enum values
UnsupportedContentType.WmfImage        // "WmfImage"
UnsupportedContentType.EmfImage        // "EmfImage"
UnsupportedContentType.SvgImage        // "SvgImage"
UnsupportedContentType.MathEquation    // "MathEquation"
UnsupportedContentType.FormField       // "FormField"
UnsupportedContentType.RubyAnnotation  // "RubyAnnotation"
UnsupportedContentType.OleObject       // "OleObject"
UnsupportedContentType.Other           // "Other"
```

## Use Cases

1. **Document Quality Assurance**: Identify documents that contain unsupported content before publishing
2. **User Feedback**: Show users what content couldn't be rendered instead of leaving gaps
3. **Debugging**: Understand why converted documents look different from the original
4. **Content Auditing**: Catalog what types of content are used across a document collection

## Limitations

- Placeholders are **visual indicators only** - they don't preserve the original content
- For math equations, consider using a library like MathJax to render the OMML content
- For form fields, the placeholder doesn't preserve interactivity
- Ruby annotations show the base text but lose the reading annotation

## Related Documentation

- [WmlToHtmlConverter Gaps](./wml_to_html_converter_gaps.md) - Full list of converter limitations
