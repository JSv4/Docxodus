# Tab Rendering Reference Documentation

These reference images were captured from LibreOffice rendering of the test files to establish the correct visual output for tab handling.

## Reference Screenshots

Generated: December 2025

### HC024-Tabs-01 and HC025-Tabs-02

Both files contain identical content demonstrating various tab types:

| Line | Tab Type | Expected Rendering |
|------|----------|-------------------|
| `left → dsfdsfsdf` | Left-aligned | Text starts at tab stop position |
| `center → bbb` | Center-aligned | "bbb" is centered AT the tab stop position |
| `right → fdssdf` | Right-aligned | "fdssdf" ends exactly at the tab stop position |
| `decimal → 1.23423` | Decimal-aligned | The decimal point aligns at the tab stop |
| `bar → dfs` | Bar tab | Vertical line at tab position, text continues |
| `left leader → ...323232` | Dot leader + left | Dots fill space, text left-aligned at stop |
| `center leader → ...dfssadf` | Dot leader + center | Dots fill space, text centered at stop |
| `right leader → ...fdsasdf` | Dot leader + right | Dots fill space, text right-aligned at stop |
| `decimal leader → ...1.23` | Dot leader + decimal | Dots fill space, decimal point at stop |
| `_________sfdajdsfjkaslfjdjkfds` | Underscore leader | Underscores fill space |
| `--------------------sfdajdsfjkaslfjdjkfds` | Hyphen leader | Hyphens fill space |

### HC026-Tabs-03

Same as above with additional content:
- `decimal leader that cau 123` - Integer without decimal point (whole number alignment)
- First line indent with tab demonstration
- Multiple leader type examples

### HC027-Tabs-04

Simple test case:
- `decimal.................123` - Decimal tab with dot leader

### HC022-Table-Of-Contents

Table of Contents with:
- "Contents" heading in blue
- TOC entries with right-aligned dot leaders to page numbers
- Format: `Video provides a powerful way to help you prove your point....................................................1`
- **Note**: Page numbers require field code resolution (separate issue)

## Tab Alignment Behavior

### Left Tab
Text starts at the tab stop position. The tab creates spacing from current position to the tab stop.

### Center Tab
Text is centered around the tab stop position. Half the text width appears before the stop, half after.

### Right Tab
Text ends at the tab stop position. The entire text appears before the stop, with trailing edge at stop.

### Decimal Tab
The decimal point (or end of integer) aligns at the tab stop position. Used for aligning numbers in columns.

### Bar Tab
A vertical line is drawn at the tab position. Text continues normally (no spacing effect).

## Leader Characters

Leaders fill the space between the current position and the tab stop with repeating characters:

| Leader Type | Character | Word XML Value |
|-------------|-----------|----------------|
| Dot | `.` | `dot` |
| Hyphen | `-` | `hyphen` |
| Underscore | `_` | `underscore` |
| Heavy | `·` (middle dot) | `heavy` |
| None | (space) | omitted or `none` |

## Current Issues (to be fixed)

1. **Leader characters not rendering** - Space instead of dots/hyphens/underscores
2. **Center alignment broken** - Renders as left-aligned
3. **Right alignment broken** - Renders as left-aligned
4. **Decimal alignment broken** - Renders as left-aligned
5. **Bar tabs not rendering** - No vertical line
6. **TOC page numbers missing** - Field codes not resolved (separate issue, documented in gaps.md)

## Root Cause Analysis

### Issue 1: Leader characters not rendering

**Location**: `WmlToHtmlConverter.cs` - `CalcWidthOfRunInTwips()` (lines 5654-5716) and `ProcessTab()` (lines 3474-3566)

**Flow**:
1. `ProcessTab()` reads `PtOpenXml.Leader` attribute (line 3479)
2. Creates a dummy run with single leader character (line 3497-3500)
3. Calls `CalcWidthOfRunInTwips(dummyRun)` to measure leader char width (line 3502)
4. `CalcWidthOfRunInTwips` returns 0 early if:
   - Font is marked unknown (line 5660-5661)
   - Font not in `KnownFamilies` (line 5670-5671)
5. Since `widthOfLeaderChar == 0`, condition at line 3515 fails
6. No leader characters are generated, falls through to empty span (line 3530-3541)

**Fix**: Modify `CalcWidthOfRunInTwips` to use estimation fallback for unknown fonts instead of returning 0, OR add special handling in `ProcessTab` for leader character measurement.

### Issue 2: Tab alignment types (center/right/decimal) not working visually

**Location**: `WmlToHtmlConverter.cs` - Tab type handling (lines 5387-5571) and `ProcessTab()` (lines 3474-3566)

**Flow**:
1. Tab type detection works correctly (lines 5387, 5431, 5514, 5557)
2. Correct `TabWidth` attribute is calculated for each type:
   - Right: `delta2 = tabPos - widthOfTextAfterTab - twipCounter` (line 5412)
   - Center: `delta2 = tabPos - (widthOfText / 2) - twipCounter` (line 5538)
   - Decimal: measures to decimal point position (lines 5449-5465)
3. `GetLeader()` is called to propagate leader attribute
4. **BUT** tab alignment type (val attribute) is NOT propagated to ProcessTab
5. `ProcessTab` only generates `margin-left` CSS for spacing
6. Text after tab still flows left-to-right regardless of tab type

**Fix**:
1. Propagate tab alignment type via new `PtOpenXml.TabAlignment` attribute
2. Modify `ProcessTab` to generate appropriate CSS:
   - Left: current behavior (margin-left)
   - Right: needs flexbox or inline-block with text-align right
   - Center: needs text-align center
   - Decimal: needs special CSS for decimal alignment

### Specific CSS Fixes Needed

**Center Tab**: Text should be centered AT the tab stop position
```css
.tab-center-container {
  display: inline-flex;
  justify-content: center;
}
```

**Right Tab**: Text should end AT the tab stop position
```css
.tab-right-container {
  display: inline-block;
  text-align: right;
}
```

**Decimal Tab**: Decimal point should align at tab stop
- Split text at decimal point
- Right-align pre-decimal portion
- Left-align post-decimal portion
