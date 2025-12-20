# WmlToHtmlConverter.cs - Gaps and Deficiencies

*Last updated: December 2025*

This document catalogs known gaps, limitations, and areas for improvement in the WmlToHtmlConverter.

## Quick Reference

| Category | Gap | Severity |
|----------|-----|----------|
| **Layout** | No `@page` CSS rule for print/PDF | Medium |
| **Layout** | Table width calculation inconsistent | Medium |
| **Layout** | Wrapper divs for simple borders | Low |
| **Layout** | Empty paragraphs verbose | Low |
| **Rendering** | Theme colors not resolved | Medium |
| **Rendering** | Text box content lost | Medium |
| **Rendering** | Tab leader count varies by platform | Low |
| **Accessibility** | No `lang` attribute on html/body | Medium |
| **Accessibility** | No `lang` attribute on foreign text spans | Medium |
| **Accessibility** | No ARIA roles | Low |
| **Fonts** | Limited font fallback (28 fonts) | Medium |
| **Fonts** | No CJK font-family fallback chain | Medium |
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

### 2. Table Width Calculation Inconsistent

**Severity:** Medium

**Problem:** Table and cell widths may not match the original document layout.

**Comparison:**
| Aspect | LibreOffice | Ours |
|--------|-------------|------|
| Units | Pixels (`width="480"`) | Points (`width: 360pt`) |
| Column widths | Proportional to content | Fixed from `tcW` |

**Impact:** Tables may appear wider or narrower than intended.

**Solution:** Review twips→points conversion in `ProcessTable` and ensure `tblW` percentage widths are handled correctly.

---

### 3. Borderless Table Detection Missing

**Severity:** Medium

**Problem:** Tables used for layout (like signature blocks) should be borderless, but borders may still appear.

**LibreOffice:**
```html
<td style="border: none; padding: 0in">
```

**Impact:** Signature blocks and multi-column layouts have unwanted borders.

**Solution:** Detect `w:tblBorders` with `w:val="nil"` or missing borders and render without CSS borders.

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

### 10. No Document Language Attribute

**Severity:** Medium

**Problem:** No `lang` attribute on `<html>` or `<body>` element.

**LibreOffice:**
```html
<body lang="en-US">
```

**Ours:** No language attribute.

**Impact:** Screen readers cannot determine document language; browsers cannot apply correct hyphenation.

**Solution:** Read from `w:settings/w:themeFontLang` or `w:lang` on document default styles.

---

### 11. No Language Attributes on Foreign Text

**Severity:** Medium

**Problem:** Text in different languages (CJK, etc.) lacks `lang` attribute.

**LibreOffice:**
```html
<span lang="zh-CN">株式会社</span>
```

**Ours:** No `lang` attribute on runs with different language.

**Impact:** Screen readers mispronounce foreign text; browsers use wrong fonts.

**Solution:** Read `w:rPr/w:lang` attributes and add `lang` to corresponding `<span>` elements.

---

### 12. No ARIA Roles

**Severity:** Low

- Images get `alt` text from `descr` attribute (good)
- No ARIA roles on semantic elements like tables, lists
- No `role="presentation"` on layout tables

---

## Font Issues

### 13. Limited Font Fallback

**Severity:** Medium
**Location:** Lines 4514-4547

Only 28 fonts have fallback definitions. Unknown fonts get no CSS `font-family` fallback to generic serif/sans-serif.

**Solution:** Add catch-all fallback: unknown fonts should fall back to `serif` or `sans-serif` based on font characteristics.

---

### 14. No CJK Font-Family Fallback Chain

**Severity:** Medium

**Problem:** CJK (Chinese, Japanese, Korean) text doesn't have proper font fallback.

**LibreOffice:**
```html
<font face="Noto Serif CJK SC">株式会社</font>
```

**Ours:** Generic serif fallback only.

**Solution:** Add CJK font-family fallback chain when CJK language detected:
```css
font-family: 'Original Font', 'Noto Serif CJK SC', 'Noto Sans CJK', 'Microsoft YaHei', 'SimSun', 'Malgun Gothic', serif;
```

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

1. **Table width calculation** - Fix twips→points conversion accuracy
2. **Borderless table detection** - For signature blocks and layout tables
3. **Theme color resolution** - Colors appear wrong with theme colors
4. **Add `lang` attribute** to `<html>` from document settings

### Medium Priority (Accessibility/Standards)

5. **Add `lang` attributes** to foreign language spans
6. **Add `@page` CSS rule** for print media
7. **CJK font-family fallback** chain
8. **Improve generic font fallback** - unknown fonts need serif/sans-serif fallback

### Low Priority (Polish)

9. **Remove wrapper divs for borders** - Apply border directly to elements
10. **Empty paragraph simplification** - Reduce HTML verbosity
11. **Line-height decimal cleanup** - `115.0%` → `115%`
12. **Render text box content** - Currently lost entirely

---

## Related Documentation

- [Unsupported Content Placeholders](./unsupported_content_placeholders.md) - Visual indicators for math, forms, WMF/EMF images
- [Comment Rendering](./comment_rendering.md) - Margin, inline, and endnote-style comment rendering
