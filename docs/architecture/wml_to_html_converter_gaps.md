# WmlToHtmlConverter.cs - Gaps and Deficiencies

*Last updated: December 2025*

This document catalogs known gaps, limitations, and areas for improvement in the WmlToHtmlConverter.

## Quick Reference

| Category | Gap | Severity |
|----------|-----|----------|
| **Layout** | No `@page` CSS rule for print/PDF | Medium |
| ~~**Layout**~~ | ~~Table width calculation inconsistent~~ | ~~Medium~~ ✅ |
| ~~**Layout**~~ | ~~Borderless table detection missing~~ | ~~Medium~~ ✅ |
| **Layout** | Wrapper divs for simple borders | Low |
| **Layout** | Empty paragraphs verbose | Low |
| **Rendering** | Theme colors not resolved | Medium |
| **Rendering** | Text box content lost | Medium |
| **Rendering** | Tab leader count varies by platform | Low |
| ~~**Accessibility**~~ | ~~No `lang` attribute on html/body~~ | ~~Medium~~ ✅ |
| ~~**Accessibility**~~ | ~~No `lang` attribute on foreign text spans~~ | ~~Medium~~ ✅ |
| **Accessibility** | No ARIA roles | Low |
| ~~**Fonts**~~ | ~~Limited font fallback (28 fonts)~~ | ~~Medium~~ ✅ |
| ~~**Fonts**~~ | ~~No CJK font-family fallback chain~~ | ~~Medium~~ ✅ |
| **Features** | Field code resolution (TOC page numbers) | Medium |
| **Features** | Pagination is CSS-only | Low |

---

## Layout Issues

### 1. No Page/Document Setup CSS

**Severity:** Medium

**Problem:** The converter does not generate `@page` CSS rules for print media or document-level settings.

**LibreOffice generates:**
```css
@page { size: 8.5in 11in; margin: 1in }
```

**Ours:** Nothing.

**Impact:** Print output and PDF generation lack proper page dimensions and margins.

**Solution:** Read page size from `w:sectPr/w:pgSz` and margins from `w:sectPr/w:pgMar`, generate `@page` CSS rule.

---

### ~~2. Table Width Calculation Inconsistent~~ ✅ RESOLVED

**Status:** Implemented in December 2025

**Solution Implemented:**
- `ProcessTable` now handles DXA widths in addition to percentage widths
- Tables with `w:tblW[@w:type="dxa"]` render with proper `width: XXpt` CSS
- Conversion formula: `dxa / 20 = points` (standard twips conversion)
- Percentage widths continue to work as before

---

### ~~3. Borderless Table Detection Missing~~ ✅ RESOLVED

**Status:** Implemented in December 2025

**Solution Implemented:**
- New `IsTableBorderless()` helper detects tables with nil/none/missing borders
- Borderless tables get `data-borderless="true"` attribute on the `<table>` element
- Checks all border sides: top, left, bottom, right, insideH, insideV
- Enables CSS-based styling: `table[data-borderless="true"] td { border: none !important; }`

---

### 4. Wrapper Divs for Simple Borders

**Severity:** Low

**Problem:** Horizontal rules are rendered with unnecessary wrapper `<div>` elements.

**LibreOffice:**
```html
<p style="border-bottom: 1px solid #cccccc">...</p>
```

**Ours:**
```html
<div class="pt-000007"><p>...</p></div>
```

**Impact:** More complex DOM, harder to style/select.

**Solution:** Apply paragraph borders directly to `<p>` element when possible.

---

### 5. Empty Paragraphs Verbose

**Severity:** Low

**Problem:** Empty paragraphs generate unnecessary markup.

**LibreOffice:**
```html
<p><br/></p>
```

**Ours:**
```html
<p dir="ltr" class="pt-Normal"><span class="pt-000000"></span></p>
```

**Impact:** Bloated HTML output.

**Solution:** Simplify empty paragraphs to `<p><br/></p>` or just `<br/>` where appropriate.

---

## Rendering Issues

### 6. Theme Colors Not Resolved

**Severity:** Medium
**Location:** Lines 3468-3472

While `w:themeColor` and `w:themeTint` are copied during border overrides, **theme colors are never resolved to actual RGB values** from the document's theme (`theme1.xml`).

**Impact:** Colors appear wrong when documents use theme colors instead of explicit RGB.

**Solution:** Read theme from `/word/theme/theme1.xml`, resolve `w:themeColor` values like `accent1`, `dark1`, etc. to RGB.

---

### 7. Text Box Content Not Fully Rendered

**Severity:** Medium
**Location:** Line 2041

Text boxes (`w:txbxContent`) are explicitly trimmed:
- `DescendantsTrimmed(W.txbxContent)` is used throughout
- Text box content is preserved but not transformed to proper HTML `<div>` or `<aside>` elements

**Impact:** Content inside text boxes is lost in HTML output.

---

### 8. Tab Leader Count Varies by Platform

**Severity:** Low

Leader character count may vary by platform due to font measurement differences:
- Desktop (.NET): Uses SkiaSharp for actual font measurement
- WASM: Uses character-based estimation

The tab span width is correct; only the dot count filling that width varies.

---

### 9. Line Height Calculation

**Severity:** Low

**Problem:** Line height values differ from LibreOffice output.

| Content Type | LibreOffice | Ours |
|--------------|-------------|------|
| Body text | `115%` | `115.0%` (extra decimal) |
| Default | `100%` | `108%` |

**Impact:** Minor spacing differences.

**Solution:** Review `w:spacing/@w:line` conversion and remove unnecessary decimal places.

---

## Accessibility Issues

### ~~10. No Document Language Attribute~~ ✅ RESOLVED

**Status:** Implemented in December 2025

**Solution Implemented:**
- `<html>` element now includes `lang` attribute (e.g., `<html lang="en-US">`)
- Language auto-detected from `w:themeFontLang` in document settings
- Falls back to default paragraph style's language, then "en-US"
- New `DocumentLanguage` setting allows manual override

---

### ~~11. No Language Attributes on Foreign Text~~ ✅ RESOLVED

**Status:** Implemented in December 2025

**Solution Implemented:**
- `GetLangAttribute()` now uses actual document default language (not hardcoded "en-US")
- Foreign text spans get `lang` attribute when different from document default
- Supports western, bidi (Arabic/Hebrew), and east Asian language detection

---

### 12. No ARIA Roles

**Severity:** Low

- Images get `alt` text from `descr` attribute (good)
- No ARIA roles on semantic elements like tables, lists
- No `role="presentation"` on layout tables

---

## Font Issues

### ~~13. Limited Font Fallback~~ ✅ RESOLVED

**Status:** Implemented in December 2025

**Solution Implemented:**
- Unknown fonts now get appropriate generic fallback (serif, sans-serif, or monospace)
- `FontClassifier` helper uses pattern matching on font names to determine fallback
- Fonts with "sans" pattern → sans-serif
- Fonts with "mono", "code", "courier" patterns → monospace
- Other fonts default to serif
- Fixed Courier New and Lucida Console to include monospace fallback

---

### ~~14. No CJK Font-Family Fallback Chain~~ ✅ RESOLVED

**Status:** Implemented in December 2025

**Solution Implemented:**
- CJK text now gets language-specific font fallback chains
- Japanese (ja-JP): `'Noto Serif CJK JP', 'Noto Sans CJK JP', 'Yu Mincho', 'MS Mincho', ...`
- Simplified Chinese (zh-hans): `'Noto Serif CJK SC', 'Microsoft YaHei', 'SimSun', ...`
- Traditional Chinese (zh-hant): `'Noto Serif CJK TC', 'Microsoft JhengHei', 'PMingLiU', ...`
- Korean (ko): `'Noto Serif CJK KR', 'Malgun Gothic', 'Batang', ...`
- Generic CJK fallback for unknown East Asian languages

---

## Feature Gaps

### 15. Field Code Resolution Not Implemented

**Severity:** Medium
**Location:** Various field handling in `ConvertToHtmlTransform`

Word documents use field codes for dynamic content. Current behavior:
- Field instructions are ignored
- Only the cached result (text between `separate` and `end`) is rendered

**Problematic scenarios:**
1. **TOC page numbers** - Empty when document was never printed/updated
2. **Cross-references** - May show stale text
3. **PAGE/NUMPAGES fields** - Cannot be resolved
4. **HYPERLINK fields** - Could be converted to `<a>` tags

---

### 16. Pagination Mode Limitations

**Severity:** Low
**Location:** `WmlToHtmlConverterSettings.PaginationMode`

The `PaginationMode.Paginated` setting is CSS-only:
- No actual page breaking logic
- Headers/footers not cloned per-page
- No page number calculation

---

### 17. Section Break Handling

**Severity:** Low
**Location:** Lines 2461-2500

- Section breaks are conflated if formatting is identical
- No visual separation or page-break CSS added
- Headers/footers for different sections not differentiated

---

## Other Issues

### 18. Missing Element Handling

Several Word elements are not handled in `ConvertToHtmlTransform`:

| Element | Purpose | Impact |
|---------|---------|--------|
| `W.softHyphen` | Soft hyphen | Loses word-wrap hints |
| `W.yearLong`, `W.monthLong`, etc. | Date/time fields | No output |
| `W.pgNum` | Page numbers | Cannot resolve |
| `W.separator`, `W.continuationSeparator` | Footnote separators | Ignored |

---

### 19. Incomplete Run Properties

**Location:** Lines 2835-2865

Unsupported run properties:
- `em` (emphasis mark)
- `emboss`, `imprint`
- `fitText`
- `kern` (kerning)
- `outline`, `shadow`
- `w` (character width expansion)

---

### 20. Paragraph Properties Not Handled

**Location:** Lines 2508-2555

Unimplemented paragraph properties:
- `framePr` (frames)
- `keepLines`, `keepNext` (pagination control)
- `mirrorIndents`
- `pageBreakBefore`
- `suppressAutoHyphens`
- `textDirection`
- `widowControl`

---

### 21. Complex Script (BiDi) Handling Incomplete

- RTL marks added but complex script font sizing (`w:szCs`) only used when `languageType == "bidi"`
- No proper handling of mixed LTR/RTL content in tables

---

### 22. Text Content in Shapes/DrawingML

Content inside DrawingML shapes (`a:txBody`, `wps:txbx`) may not be fully extracted:
- Shape text handled differently than regular paragraph text
- Nested text frames in complex drawings may be missed
- No CSS positioning to reflect shape placement

---

### 23. Tracked Changes - Partial Property Support

**Location:** Lines 3003-3054

`DescribeFormatChange` only checks 7 properties:
- Bold, Italic, Underline, Strikethrough, Font size, Font name, Color

Missing: highlight, caps, smallCaps, spacing, position, etc.

---

## Summary of Priority Fixes

### High Priority (Visual/Layout Impact)

1. ~~**Table width calculation** - Fix twips→points conversion accuracy~~ ✅ Done
2. ~~**Borderless table detection** - For signature blocks and layout tables~~ ✅ Done
3. **Theme color resolution** - Colors appear wrong with theme colors
4. ~~**Add `lang` attribute** to `<html>` from document settings~~ ✅ Done

### Medium Priority (Accessibility/Standards)

5. ~~**Add `lang` attributes** to foreign language spans~~ ✅ Done
6. **Add `@page` CSS rule** for print media
7. ~~**CJK font-family fallback** chain~~ ✅ Done
8. ~~**Improve generic font fallback** - unknown fonts need serif/sans-serif fallback~~ ✅ Done

### Low Priority (Polish)

9. **Remove wrapper divs for borders** - Apply border directly to elements
10. **Empty paragraph simplification** - Reduce HTML verbosity
11. **Line-height decimal cleanup** - `115.0%` → `115%`
12. **Render text box content** - Currently lost entirely

---

## Related Documentation

- [Unsupported Content Placeholders](./unsupported_content_placeholders.md) - Visual indicators for math, forms, WMF/EMF images
- [Comment Rendering](./comment_rendering.md) - Margin, inline, and endnote-style comment rendering
