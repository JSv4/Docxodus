# WmlToHtmlConverter.cs - Gaps and Deficiencies

*Last updated: December 2025*

This document catalogs known gaps, limitations, and areas for improvement in the WmlToHtmlConverter.

## Quick Reference

| Category | Gap | Severity |
|----------|-----|----------|
| ~~**Stability**~~ | ~~Null reference crashes in `DefineRunStyle`~~ | ~~High~~ FIXED |
| ~~**Stability**~~ | ~~Static caches not thread-safe~~ | ~~High~~ FIXED |
| ~~**Rendering**~~ | ~~Tab width calculation disabled~~ | ~~High~~ FIXED |
| **Rendering** | Theme colors not resolved | Medium |
| **Rendering** | Text box content lost | Medium |
| **Rendering** | SVG images not supported | Medium |
| **Rendering** | WMF/EMF images excluded | Low |
| **Rendering** | Tab leader count varies by platform | Low |
| **Features** | Field code resolution (TOC page numbers) | Medium |
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

## 7. ~~Tab Width Calculation Disabled~~ (FIXED)

**Status:** Resolved

Previously, tab width calculation for text elements was disabled with `const int widthOfText = 0;` at line 5591. This was disabled because "it doesn't work on Azure" (likely due to font unavailability).

Now uses estimation fallback when font measurement fails:
- `MetricsGetter._getTextWidth()` returns character-based estimation when SkiaSharp measurement fails
- Estimation formula: `charWidth = fontSize * 0.6 / 2` per character
- Works in Azure, WASM, and environments without fonts installed
- Tab positioning now properly accounts for preceding text width
- Leader spans now have `display: inline-block` for proper width rendering

**Note:** Leader character count may vary by platform due to font measurement differences:
- Desktop (.NET): Uses SkiaSharp for actual font measurement - period characters may measure wider than expected, resulting in fewer dots
- WASM: Uses character-based estimation - may produce different counts
- The tab span width is correct; only the dot count filling that width varies

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

## 12. Field Code Resolution Not Implemented

**Location:** Various field handling in `ConvertToHtmlTransform`

Word documents use field codes for dynamic content like page numbers, table of contents, cross-references, and calculated values. These are stored as:
- `w:fldSimple` - Simple fields with direct content
- `w:fldChar` with `w:fldCharType="begin"/"separate"/"end"` - Complex fields with instruction and result parts
- `w:instrText` - Field instruction text (e.g., `PAGE`, `TOC`, `HYPERLINK`)

**Current behavior:**
- Field instructions are ignored
- Only the cached result (text between `separate` and `end`) is rendered
- This works for static document snapshots but fails for:

**Problematic scenarios:**
1. **TOC page numbers** - Appear as `#x200e` (Unicode LRM) because the cached result is empty when document was never printed/updated
2. **Cross-references** - May show stale or placeholder text
3. **PAGE/NUMPAGES fields** - Cannot be resolved (would require actual pagination)
4. **Calculated fields** - Results may be outdated

**Example from Table of Contents:**
```xml
<w:fldSimple w:instr=" PAGEREF _Toc123 \h ">
  <w:r><w:t>3</w:t></w:r>
</w:fldSimple>
```
If the cached result is empty (common when document hasn't been printed), the page number simply doesn't appear.

**Potential solutions:**
1. **Warn when fields have no cached result** - Emit visible placeholder or console warning
2. **Parse simple field types** - Resolve `HYPERLINK` fields to actual `<a>` tags
3. **TOC-specific handling** - Detect TOC fields and warn about missing page numbers
4. **Full field code parsing** - Complex; would require understanding all field types

## 13. Tracked Changes - Partial Property Support

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
4. ~~**Fix tab width calculation**~~ - FIXED - now uses estimation fallback when fonts unavailable

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
