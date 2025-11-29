# WmlToHtmlConverter.cs - Gaps and Deficiencies

This document catalogs known gaps, limitations, and areas for improvement in the WmlToHtmlConverter.

## 1. Missing Element Handling

Several Word elements are not handled in `ConvertToHtmlTransform`:

| Element | Purpose | Impact |
|---------|---------|--------|
| `W.softHyphen` | Soft hyphen | Silently ignored, loses word-wrap hints |
| `W.yearLong`, `W.yearShort`, `W.monthLong`, `W.monthShort`, `W.dayLong`, `W.dayShort` | Date/time fields | No output |
| `W.pgNum`, `W.fldChar`, `W.fldSimple` (partially) | Page numbers, field codes | Incomplete field support |
| `W.ruby` | Ruby annotations (CJK) | No support for East Asian text annotations |
| `W.separator`, `W.continuationSeparator` | Footnote separators | Ignored |

## 2. ~~Comment Rendering Mode Incomplete~~ (FIXED)

**Status:** Resolved

`CommentRenderMode.Margin` is now fully implemented with:
- Flexbox-based layout with main content and margin column
- CSS styling for margin notes with author, date, and back-reference links
- Print media query for responsive behavior

See `docs/architecture/comment_rendering.md` for full documentation.

## 3. Text Box Content Not Fully Rendered

**Location:** Line 2041

Text boxes (`w:txbxContent`) are explicitly trimmed in multiple places:
- `DescendantsTrimmed(W.txbxContent)` is used throughout
- Text box content is preserved but not transformed to proper HTML `<div>` or `<aside>` elements

## 4. Limited Drawing/Image Support

**Location:** Lines 4617-4856

- Only handles these content types: `png`, `gif`, `tiff`, `jpeg`
- **WMF and EMF files are explicitly excluded** (line 4619 comment)
- SVG images not supported
- **No fallback or placeholder** for unsupported image types - they just disappear

## 5. Incomplete Run Properties

**Location:** Lines 2835-2865

The code documents unsupported properties:
```csharp
// Don't handle:
// - em (emphasis mark)
// - emboss
// - fitText
// - imprint
// - kern (kerning)
// - outline
// - shadow
// - w (character width expansion)
```

## 6. Paragraph Properties Not Handled

**Location:** Lines 2508-2555

Many documented but unimplemented:
- `contextualSpacing` - partially handled
- `framePr` (frames)
- `keepLines`, `keepNext` (pagination control)
- `mirrorIndents`
- `pageBreakBefore`
- `suppressAutoHyphens`
- `tabs` - only partially implemented
- `textDirection`
- `widowControl`

## 7. Tab Width Calculation Disabled

**Location:** Lines 3803-3823

```csharp
// TODO: Revisit. This is a quick fix because it doesn't work on Azure.
// ...
const int widthOfText = 0;  // <-- Always zero!
```

Tab width calculation for text content is completely disabled.

## 8. Hard-coded Default Language

**Location:** Line 3211

```csharp
const string defaultLanguage = "en-US"; // todo need to get defaultLanguage
```

Should read from document settings (`w:settings/w:themeFontLang`).

## 9. Theme Colors Not Resolved

**Location:** Lines 3468-3472

While `w:themeColor` and `w:themeTint` are copied during border overrides, **theme colors are never resolved to actual RGB values** from the document's theme.

## 10. No Math Equation Support

OMML (`<m:oMath>`) elements are not handled at all - equations silently disappear from output.

## 11. Section Break Handling

**Location:** Lines 2461-2500

- Section breaks are conflated if formatting is identical
- No visual separation or page-break CSS added
- Headers/footers for different sections not differentiated

## 12. Tracked Changes - Partial Property Support

**Location:** Lines 3003-3054

`DescribeFormatChange` only checks 7 properties:
- Bold, Italic, Underline, Strikethrough, Font size, Font name, Color

Missing: highlight, caps, smallCaps, spacing, position, etc.

## 13. Potential Null Reference Issues

Several places access `.First()` without null checks:
- Line 3061: `var rPr = run.Elements(W.rPr).First();` - crashes if no `rPr`
- Line 3213: `var rPr = run.Elements(W.rPr).First();`

## 14. Font Fallback Limited

**Location:** Lines 4514-4547

Only 28 fonts have fallback definitions. Unknown fonts get no CSS `font-family` fallback to serif/sans-serif.

## 15. Static Mutable State

**Location:** Lines 3883, 4410

```csharp
private static readonly HashSet<string> UnknownFonts = new HashSet<string>();
private static readonly Dictionary<string, string> ShadeCache = new Dictionary<string, string>();
```

These are not thread-safe and will grow unbounded across multiple document conversions.

## 16. Complex Script (BiDi) Handling Incomplete

- RTL marks added but complex script font sizing (`w:szCs`) is only used when `languageType == "bidi"`
- No proper handling of mixed LTR/RTL content in tables

## 17. No Accessibility Attributes

- Images get `alt` text from `descr` attribute
- No ARIA roles on semantic elements
- No `lang` attribute on the `<html>` element itself

## 18. Form Fields Not Supported

`w:ffData`, `w:checkBox`, `w:textInput`, `w:ddList` are not converted to HTML form elements.

---

## Summary of Priority Fixes

### High Priority

1. ~~**Implement `CommentRenderMode.Margin`**~~ - FIXED
2. **Handle null `rPr`** in `DefineRunStyle` and `GetLangAttribute` to prevent crashes
3. **Add thread-safety** to static caches or make them instance-based

### Medium Priority

4. **Add SVG image support** - increasingly common in modern documents
5. **Implement theme color resolution** for accurate color rendering
6. **Fix tab width calculation** - currently disabled entirely

### Low Priority (Feature Additions)

7. **Consider OMML to MathML conversion** for equation support
8. **Add form field support** for interactive documents
9. **Improve accessibility** with ARIA roles and proper `lang` attributes
