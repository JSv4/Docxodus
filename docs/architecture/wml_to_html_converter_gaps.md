# WmlToHtmlConverter.cs - Gaps and Deficiencies

*Last updated: December 2025*

This document catalogs known gaps, limitations, and areas for improvement in the WmlToHtmlConverter.

## Quick Reference

| Category | Gap | Severity |
|----------|-----|----------|
| ~~**Stability**~~ | ~~Null reference crashes in `DefineRunStyle`~~ | ~~High~~ FIXED |
| ~~**Stability**~~ | ~~Static caches not thread-safe~~ | ~~High~~ FIXED |
| **Rendering** | Tab width calculation disabled | High |
| **Rendering** | Theme colors not resolved | Medium |
| **Rendering** | Text box content lost | Medium |
| **Rendering** | SVG images not supported | Medium |
| **Rendering** | WMF/EMF images excluded | Low |
| **Features** | Math equations (OMML) not rendered | Medium |
| **Features** | Form fields not supported | Low |
| **Features** | Pagination is CSS-only | Low |
| **Accessibility** | No ARIA roles or lang attribute | Low |

---

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

## 13. ~~Potential Null Reference Issues~~ (FIXED)

**Status:** Resolved

Previously, `DefineRunStyle` and `GetLangAttribute` used `.First()` on `run.Elements(W.rPr)` which would crash with `InvalidOperationException` if a run had no `rPr` element. Now uses `.FirstOrDefault()` with null checks to return gracefully (empty style dictionary or null language attribute).

## 14. Font Fallback Limited

**Location:** Lines 4514-4547

Only 28 fonts have fallback definitions. Unknown fonts get no CSS `font-family` fallback to serif/sans-serif.

## 15. ~~Static Mutable State~~ (FIXED)

**Status:** Resolved

Previously, static caches were not thread-safe:
- `UnknownFonts` in `FontFamilyHelper.cs`
- `ShadeCache` in `WmlToHtmlConverter.cs`

Now uses thread-safe `ConcurrentDictionary` for both caches, with `Lazy<T>` for font family initialization. Added `ClearShadeCache()` and `ClearUnknownFontsCache()` methods for memory management in long-running processes.

## 16. Complex Script (BiDi) Handling Incomplete

- RTL marks added but complex script font sizing (`w:szCs`) is only used when `languageType == "bidi"`
- No proper handling of mixed LTR/RTL content in tables

## 17. No Accessibility Attributes

- Images get `alt` text from `descr` attribute
- No ARIA roles on semantic elements
- No `lang` attribute on the `<html>` element itself

## 18. Form Fields Not Supported

`w:ffData`, `w:checkBox`, `w:textInput`, `w:ddList` are not converted to HTML form elements.

## 19. Pagination Mode Limitations

**Location:** `WmlToHtmlConverterSettings.PaginationMode`

The `PaginationMode.Paginated` setting is architecturally implemented but has significant limitations:

- **CSS only** - Generates PDF.js-style styling but content still flows continuously
- **No actual page breaking** - No page-break logic or layout engine
- **Headers/footers must be cloned per-page** - Dynamic fields like PAGE number don't work
- **Section boundaries not detected** - Pagination engine doesn't track which section a page belongs to
- **No page number calculation** - Cannot determine total page count

This is essentially a styling mode rather than true pagination.

## 20. Text Content in Shapes/DrawingML

Beyond text boxes, content inside DrawingML shapes (`a:txBody`, `wps:txbx`) may not be fully extracted:

- Shape text is handled differently than regular paragraph text
- Nested text frames in complex drawings may be missed
- No CSS positioning to reflect shape placement

---

## Summary of Priority Fixes

### High Priority (Stability/Correctness)

1. ~~**Implement `CommentRenderMode.Margin`**~~ - FIXED
2. ~~**Handle null `rPr`** in `DefineRunStyle` and `GetLangAttribute` to prevent crashes~~ - FIXED
3. ~~**Add thread-safety** to static caches or make them instance-based (memory leak in high-volume scenarios)~~ - FIXED
4. **Fix tab width calculation** - currently hardcoded to 0, making tabulated content unreadable

### Medium Priority (Visual Fidelity)

5. **Implement theme color resolution** - colors appear wrong when documents use theme colors
6. **Add SVG image support** - increasingly common in modern documents
7. **Render text box content** - currently lost entirely from output
8. **Improve font fallback** - unknown fonts should fall back to generic serif/sans-serif

### Low Priority (Feature Additions)

9. **Consider OMML to MathML conversion** for equation support
10. **Add form field support** for interactive documents
11. **Improve accessibility** with ARIA roles and proper `lang` attributes
12. **Add WMF/EMF conversion** or placeholder rendering for legacy images
