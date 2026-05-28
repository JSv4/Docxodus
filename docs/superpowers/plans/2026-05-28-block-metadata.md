# Block Metadata Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Expose a read-only block-metadata surface on `DocxSession` (`GetBlockMetadata`, `GetBlockMetadatas`, `GetListMembership`, `GetSectionInfo`) so callers can introspect a paragraph's style id, style name, outline level, numbering identity (numId/abstractNumId/ilvl/format/start override/inherited-from-style flag), and the enclosing `w:sectPr` (page size/orientation/margins/columns/header/footer parts) without falling back to `Raw.GetXml`.

**Architecture:** Three new public methods on `DocxSession` (plus the bulk variant of `GetBlockMetadata`). Implementation factored into `Internal/BlockMetadataOps.cs` (pure read; no mutation, no undo snapshot) — mirrors the existing `Internal/AnnotationOps.cs` factoring. Reuses `FindAnchor` / `Resolve` infrastructure already in `DocxSession`. Wires through the standard ripple: `DocxSessionOps` → `DocxSessionJson` → WASM `[JSExport]` shell → stdio dispatcher → npm wrapper → Python wrapper. Tests cover every record field + null paths.

**Tech Stack:** .NET 8 / Open XML SDK 3.x / xUnit / TypeScript / Playwright / Python 3.10+ / pytest.

**Spec:** `docs/superpowers/specs/2026-05-28-note-refs-list-writes-metadata-design.md`

---

## File Structure

**New files:**
- `Docxodus/Internal/BlockMetadataOps.cs` — resolves `BlockMetadata` / `ListMembership` / `SectionInfo` from a `WordprocessingDocument` + anchor target
- `Docxodus.Tests/DocxSessionMetadataTests.cs` — xUnit coverage
- `npm/tests/block-metadata.spec.ts` — Playwright spec
- `python/tests/test_block_metadata.py` — pytest spec

**Modified files:**
- `Docxodus/DocxSession.cs` — add `NumberFormat` enum + `BlockMetadata` / `ListMembership` / `SectionInfo` records; add the four public methods
- `Docxodus/Internal/DocxSessionOps.cs` — add four facade methods
- `Docxodus/Internal/DocxSessionJson.cs` — add four serializers (`SerializeBlockMetadataOrNull`, `SerializeBlockMetadataMap`, `SerializeListMembershipOrNull`, `SerializeSectionInfoOrNull`) and a string helper for the `NumberFormat` enum
- `wasm/DocxodusWasm/DocxSessionBridge.cs` — add four `[JSExport]` shells
- `tools/python-host/Dispatcher.cs` — add four switch cases
- `npm/src/types.ts` — add `NumberFormat` const enum + four interfaces + four `DocxodusWasmExports` signatures
- `npm/src/session.ts` — add four methods on the `DocxSession` class
- `npm/src/index.ts` — re-export the new types
- `python/src/docx_scalpel/types.py` — add four dataclasses + `NumberFormat` enum
- `python/src/docx_scalpel/session.py` — add four methods on the `DocxSession` class
- `CHANGELOG.md` — Unreleased entry
- `CLAUDE.md` — extend the DocxSession bullet list
- `docs/architecture/docx_mutation_api.md` — new "Inspection: block metadata" section

---

## Phase 1: Core types and read helpers

### Task 1: Add the new public types to `DocxSession.cs`

**Files:**
- Modify: `Docxodus/DocxSession.cs` (insert types alongside existing public records — after `AnchorInfo` at line 451)

The plan defers all method additions to later tasks. This task only adds the types so subsequent tasks can compile.

- [ ] **Step 1: Read the file to locate the insertion point**

```bash
grep -n "public sealed record AnchorInfo" Docxodus/DocxSession.cs
```
Expected: one hit, line 434 (`public sealed record AnchorInfo(...)`).

- [ ] **Step 2: Insert the new types directly after the closing `}` of the `AnchorInfo` record (the spot right before `public sealed record EditSummary`)**

```csharp
/// <summary>
/// The six list formats supported by the list write surface
/// (<c>InsertNumberedList</c>, <c>ConvertToNumberedList</c>, …) and
/// surfaced on <see cref="ListMembership.Format"/>. Maps to OOXML
/// <c>w:numFmt</c> values: <c>Decimal</c> → <c>decimal</c>,
/// <c>UpperLetter</c> → <c>upperLetter</c>, <c>LowerLetter</c> →
/// <c>lowerLetter</c>, <c>UpperRoman</c> → <c>upperRoman</c>,
/// <c>LowerRoman</c> → <c>lowerRoman</c>, <c>Bullet</c> → <c>bullet</c>.
/// Other OOXML formats resolve to <c>Decimal</c> (the safest fallback).
/// </summary>
public enum NumberFormat
{
    Decimal,
    UpperLetter,
    LowerLetter,
    UpperRoman,
    LowerRoman,
    Bullet,
}

/// <summary>
/// Numbering facts for a list-item paragraph. Returned by
/// <see cref="DocxSession.GetListMembership"/> and surfaced as
/// <see cref="BlockMetadata.List"/>.
/// </summary>
public sealed record ListMembership
{
    /// <summary>The <c>w:numId</c> the paragraph belongs to (the <c>w:num</c> instance).</summary>
    required public int NumId { get; init; }

    /// <summary>The <c>w:abstractNumId</c> the paragraph's <c>w:num</c> points at (the format template).</summary>
    required public int AbstractNumId { get; init; }

    /// <summary>The paragraph's level (<c>w:ilvl</c>), 0-8.</summary>
    required public int Level { get; init; }

    /// <summary>The resolved <see cref="NumberFormat"/> for this paragraph's level.</summary>
    required public NumberFormat Format { get; init; }

    /// <summary>The start-override applied to this paragraph's level via
    /// <c>w:lvlOverride/w:startOverride</c>, if any. <c>null</c> when no override is in effect.</summary>
    public int? StartOverride { get; init; }

    /// <summary>Always <c>true</c> for a paragraph carrying <c>w:numPr</c> (inline or via style).</summary>
    required public bool IsAutoNumbered { get; init; }

    /// <summary><c>true</c> when the <c>w:numPr</c> is inherited from the paragraph style chain
    /// rather than set directly on the paragraph. <c>false</c> when set inline on the paragraph.</summary>
    required public bool FromStyle { get; init; }

    /// <summary>The rendered auto-number prefix (e.g. <c>"1."</c>, <c>"(a)"</c>) — same value
    /// surfaced as <see cref="AnchorInfo.AutoNumberPrefix"/>. Duplicated here so callers don't
    /// have to take two round-trips.</summary>
    public string? GeneratedLabel { get; init; }
}

/// <summary>
/// Block-level structural metadata. Returned by <see cref="DocxSession.GetBlockMetadata"/>.
/// </summary>
public sealed record BlockMetadata
{
    required public string AnchorId { get; init; }
    required public string Kind { get; init; }
    required public string Scope { get; init; }

    /// <summary>The <c>w:pStyle/@w:val</c> for paragraph kinds, or <c>w:tblStyle</c> for tables.
    /// <c>null</c> when no style is applied.</summary>
    public string? StyleId { get; init; }

    /// <summary>Resolved <c>w:name/@w:val</c> for <see cref="StyleId"/> from the styles part.
    /// <c>null</c> when styles part is absent or the style isn't defined.</summary>
    public string? StyleName { get; init; }

    /// <summary>Outline level: <c>w:pPr/w:outlineLvl</c> when present; otherwise
    /// inferred from a Heading1..Heading9 style (level 0..8); <c>null</c> otherwise.
    /// Word's outlineLvl is 0-based (0 = top heading).</summary>
    public int? OutlineLevel { get; init; }

    /// <summary>Populated for list-item paragraphs; <c>null</c> otherwise.</summary>
    public ListMembership? List { get; init; }

    /// <summary><c>true</c> when any descendant <c>w:r</c> carries a non-empty <c>w:rPr</c>
    /// (bold, italic, color, run style, etc.). Coarse but useful as a "does this paragraph
    /// have inline formatting at all?" probe.</summary>
    required public bool HasInlineFormatting { get; init; }
}

/// <summary>
/// Page-layout snapshot for the <c>w:sectPr</c> that governs an anchor.
/// Returned by <see cref="DocxSession.GetSectionInfo"/>; <c>null</c> for
/// anchors outside the body part (footnotes/endnotes/headers/footers/comments).
/// </summary>
public sealed record SectionInfo
{
    /// <summary>The Unid of the <c>w:sectPr</c> element this info describes. Stable across mutations.</summary>
    required public string SectionUnid { get; init; }

    required public double PageWidthTwips { get; init; }
    required public double PageHeightTwips { get; init; }
    required public bool Landscape { get; init; }
    required public double MarginTopTwips { get; init; }
    required public double MarginBottomTwips { get; init; }
    required public double MarginLeftTwips { get; init; }
    required public double MarginRightTwips { get; init; }

    /// <summary>Number of text columns. Defaults to 1 if no <c>w:cols</c> is set.</summary>
    required public int Columns { get; init; }

    /// <summary>URIs of the header parts referenced by this section, in declaration order.</summary>
    required public IReadOnlyList<string> HeaderPartUris { get; init; }

    /// <summary>URIs of the footer parts referenced by this section, in declaration order.</summary>
    required public IReadOnlyList<string> FooterPartUris { get; init; }
}

```

- [ ] **Step 3: Build to verify types compile**

```bash
dotnet build Docxodus/Docxodus.csproj
```
Expected: build succeeds.

- [ ] **Step 4: Commit**

```bash
git add Docxodus/DocxSession.cs
git commit -m "feat(session): add BlockMetadata, ListMembership, SectionInfo, NumberFormat types"
```

---

### Task 2: Write the failing test fixture for `BlockMetadataOps`

**Files:**
- Create: `Docxodus.Tests/DocxSessionMetadataTests.cs`

We start with a tiny test that exercises the simplest case (`GetBlockMetadata` on a plain paragraph). Subsequent tasks expand coverage.

- [ ] **Step 1: Create the file with one failing test**

```csharp
#nullable enable

// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Linq;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Tests for the block-metadata read surface on <see cref="DocxSession"/>
/// (<c>GetBlockMetadata</c>, <c>GetBlockMetadatas</c>, <c>GetListMembership</c>,
/// <c>GetSectionInfo</c>). Test IDs follow the <c>BM###</c> prefix convention.
/// </summary>
public class DocxSessionMetadataTests
{
    [Fact]
    public void BM001_GetBlockMetadata_PlainParagraph_ReturnsKindAndScope()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = session.Project().AnchorIndex.Values.First(t => t.Anchor.Kind == "p");

        var meta = session.GetBlockMetadata(anchor.Anchor.Id);

        Assert.NotNull(meta);
        Assert.Equal("p", meta!.Kind);
        Assert.Equal("body", meta.Scope);
        Assert.Null(meta.StyleId);
        Assert.Null(meta.StyleName);
        Assert.Null(meta.OutlineLevel);
        Assert.Null(meta.List);
        Assert.False(meta.HasInlineFormatting);
    }
}
```

- [ ] **Step 2: Run to verify it fails (method not defined)**

```bash
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~BM001"
```
Expected: BUILD FAIL with `CS1061: 'DocxSession' does not contain a definition for 'GetBlockMetadata'`.

- [ ] **Step 3: Commit the failing test**

```bash
git add Docxodus.Tests/DocxSessionMetadataTests.cs
git commit -m "test(session): BM001 failing test for GetBlockMetadata on plain paragraph"
```

---

### Task 3: Create `Internal/BlockMetadataOps.cs` with the resolver skeleton

**Files:**
- Create: `Docxodus/Internal/BlockMetadataOps.cs`

Pure functions over a `WordprocessingDocument` + `AnchorTarget` — no state, no undo, no mutation.

- [ ] **Step 1: Create the file with the resolver skeleton**

```csharp
#nullable enable

// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus.Internal;

/// <summary>
/// Pure resolvers for the block-metadata read surface — no mutation, no undo
/// snapshots. Each method takes a live <see cref="WordprocessingDocument"/>
/// and an <see cref="AnchorTarget"/> and walks the OOXML tree to assemble
/// the requested record.
/// </summary>
internal static class BlockMetadataOps
{
    /// <summary>Resolve <see cref="BlockMetadata"/> for the given anchor target.</summary>
    public static BlockMetadata? GetBlockMetadata(WordprocessingDocument doc, AnchorTarget target)
    {
        var element = target.Resolve(doc);
        if (element is null) return null;

        var styleId = ResolveStyleId(element);
        var styleName = styleId is null ? null : ResolveStyleName(doc, styleId);
        var outlineLevel = ResolveOutlineLevel(element, styleId);
        var list = ResolveListMembership(doc, element, target);
        var hasFormatting = HasInlineFormatting(element);

        return new BlockMetadata
        {
            AnchorId = target.Anchor.Id,
            Kind = target.Anchor.Kind,
            Scope = target.Anchor.Scope,
            StyleId = styleId,
            StyleName = styleName,
            OutlineLevel = outlineLevel,
            List = list,
            HasInlineFormatting = hasFormatting,
        };
    }

    /// <summary>Resolve <see cref="ListMembership"/> for the given target, or null if not a list item.</summary>
    public static ListMembership? GetListMembership(WordprocessingDocument doc, AnchorTarget target)
    {
        var element = target.Resolve(doc);
        return element is null ? null : ResolveListMembership(doc, element, target);
    }

    /// <summary>Resolve <see cref="SectionInfo"/> for the given target, or null for non-body anchors.</summary>
    public static SectionInfo? GetSectionInfo(WordprocessingDocument doc, AnchorTarget target)
    {
        // Implemented in Task 7.
        return null;
    }

    // ─── Internals ──────────────────────────────────────────────────────

    private static string? ResolveStyleId(XElement element)
    {
        if (element.Name == W.p)
            return (string?)element.Element(W.pPr)?.Element(W.pStyle)?.Attribute(W.val);
        if (element.Name == W.tbl)
            return (string?)element.Element(W.tblPr)?.Element(W.tblStyle)?.Attribute(W.val);
        return null;
    }

    private static string? ResolveStyleName(WordprocessingDocument doc, string styleId)
    {
        var stylesPart = doc.MainDocumentPart?.StyleDefinitionsPart;
        var root = stylesPart?.GetXDocument().Root;
        if (root is null) return null;
        var style = root.Elements(W.style).FirstOrDefault(s => (string?)s.Attribute(W.styleId) == styleId);
        return (string?)style?.Element(W.name)?.Attribute(W.val);
    }

    private static int? ResolveOutlineLevel(XElement element, string? styleId)
    {
        if (element.Name != W.p) return null;

        var outlineLvl = element.Element(W.pPr)?.Element(W.outlineLvl);
        if (outlineLvl is not null && int.TryParse((string?)outlineLvl.Attribute(W.val), out var lvl))
            return lvl;

        // Infer from Heading1..Heading9 style id.
        if (styleId is { Length: > 7 }
            && styleId.StartsWith("Heading", System.StringComparison.Ordinal)
            && int.TryParse(styleId.AsSpan(7), out var hLvl)
            && hLvl >= 1 && hLvl <= 9)
        {
            return hLvl - 1;  // outlineLvl is 0-based
        }
        return null;
    }

    private static ListMembership? ResolveListMembership(
        WordprocessingDocument doc, XElement element, AnchorTarget target)
    {
        // Implemented in Task 5.
        return null;
    }

    private static bool HasInlineFormatting(XElement element)
    {
        return element.Descendants(W.r).Any(r =>
        {
            var rPr = r.Element(W.rPr);
            return rPr is not null && rPr.Elements().Any();
        });
    }
}
```

- [ ] **Step 2: Verify the build (BlockMetadataOps compiles even though `DocxSession` doesn't call it yet)**

```bash
dotnet build Docxodus/Docxodus.csproj
```
Expected: build succeeds.

- [ ] **Step 3: Commit the skeleton**

```bash
git add Docxodus/Internal/BlockMetadataOps.cs
git commit -m "feat(session): BlockMetadataOps skeleton (style id/name, outline level, formatting probe)"
```

---

### Task 4: Wire `DocxSession.GetBlockMetadata` and make BM001 pass

**Files:**
- Modify: `Docxodus/DocxSession.cs`

Add the public method that calls `BlockMetadataOps.GetBlockMetadata` after the existing `GetAnchorInfos` method (line 914 area).

- [ ] **Step 1: Locate the insertion point**

```bash
grep -n "public IReadOnlyDictionary<string, AnchorInfo?> GetAnchorInfos" Docxodus/DocxSession.cs
```
Expected: one hit around line 894.

- [ ] **Step 2: Insert the new method right after `GetAnchorInfos`'s closing `}` (the line with just `    }` before `/// <summary>` for `Grep`)**

```csharp
    /// <summary>
    /// Resolves block-level metadata (style id + name, outline level, list
    /// membership, formatting probe) for <paramref name="anchorId"/>. Returns
    /// <c>null</c> when the anchor doesn't exist. See <see cref="BlockMetadata"/>
    /// for the field reference.
    /// </summary>
    public BlockMetadata? GetBlockMetadata(string anchorId)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(anchorId);
        var target = FindAnchor(anchorId);
        return target is null ? null : Internal.BlockMetadataOps.GetBlockMetadata(_doc!, target);
    }

    /// <summary>
    /// Bulk variant of <see cref="GetBlockMetadata"/>. Unknown anchor ids map
    /// to <c>null</c>; duplicate ids are deduped; iteration order matches
    /// input order for keys that appear first.
    /// </summary>
    public IReadOnlyDictionary<string, BlockMetadata?> GetBlockMetadatas(IEnumerable<string> anchorIds)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(anchorIds);

        var result = new Dictionary<string, BlockMetadata?>(StringComparer.Ordinal);
        foreach (var id in anchorIds)
        {
            if (id is null) continue;
            if (result.ContainsKey(id)) continue;
            var target = FindAnchor(id);
            result[id] = target is null ? null : Internal.BlockMetadataOps.GetBlockMetadata(_doc!, target);
        }
        return result;
    }

    /// <summary>
    /// Resolves the <see cref="ListMembership"/> for a list-item paragraph;
    /// returns <c>null</c> when the anchor has no <c>w:numPr</c> (inline or
    /// inherited from style) or doesn't exist.
    /// </summary>
    public ListMembership? GetListMembership(string anchorId)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(anchorId);
        var target = FindAnchor(anchorId);
        return target is null ? null : Internal.BlockMetadataOps.GetListMembership(_doc!, target);
    }

    /// <summary>
    /// Resolves the <see cref="SectionInfo"/> for the <c>w:sectPr</c> that
    /// governs <paramref name="anchorId"/>. Returns <c>null</c> when the
    /// anchor lives outside the body part (footnotes, endnotes, headers,
    /// footers, comments) or doesn't exist.
    /// </summary>
    public SectionInfo? GetSectionInfo(string anchorId)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(anchorId);
        var target = FindAnchor(anchorId);
        return target is null ? null : Internal.BlockMetadataOps.GetSectionInfo(_doc!, target);
    }

```

- [ ] **Step 3: Run BM001 to verify it passes**

```bash
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~BM001"
```
Expected: 1 passed.

- [ ] **Step 4: Commit**

```bash
git add Docxodus/DocxSession.cs
git commit -m "feat(session): add GetBlockMetadata / GetListMembership / GetSectionInfo public methods"
```

---

## Phase 2: List membership resolution

### Task 5: Add tests + implementation for `ListMembership` (inline `w:numPr`)

**Files:**
- Modify: `Docxodus.Tests/DocxSessionMetadataTests.cs`
- Modify: `Docxodus/Internal/BlockMetadataOps.cs`

- [ ] **Step 1: Add failing test BM002 to the test class**

Append inside `DocxSessionMetadataTests`:

```csharp
    [Fact]
    public void BM002_GetListMembership_InlineNumPr_BulletList_ReturnsListFacts()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS002_BulletedList());
        var anchor = session.Project().AnchorIndex.Values.First(t => t.Anchor.Kind == "li");

        var list = session.GetListMembership(anchor.Anchor.Id);

        Assert.NotNull(list);
        Assert.Equal(1, list!.NumId);
        Assert.Equal(0, list.AbstractNumId);
        Assert.Equal(0, list.Level);
        Assert.Equal(NumberFormat.Bullet, list.Format);
        Assert.True(list.IsAutoNumbered);
        Assert.False(list.FromStyle);
        Assert.Null(list.StartOverride);
    }

    [Fact]
    public void BM003_GetListMembership_NotAList_ReturnsNull()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = session.Project().AnchorIndex.Values.First(t => t.Anchor.Kind == "p");

        Assert.Null(session.GetListMembership(anchor.Anchor.Id));
    }
```

- [ ] **Step 2: Run to verify they fail**

```bash
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~BM002|FullyQualifiedName~BM003"
```
Expected: BM002 FAIL (assertion `list` is null), BM003 PASS (resolver already returns null).

- [ ] **Step 3: Implement `ResolveListMembership` in `BlockMetadataOps.cs`**

Replace the stub `ResolveListMembership` with:

```csharp
    private static ListMembership? ResolveListMembership(
        WordprocessingDocument doc, XElement element, AnchorTarget target)
    {
        if (element.Name != W.p) return null;

        // Inline w:numPr beats style-inherited numPr.
        var pPr = element.Element(W.pPr);
        var numPr = pPr?.Element(W.numPr);
        bool fromStyle = false;

        if (numPr is null)
        {
            // Walk the pStyle chain looking for a style that contributes numPr.
            var styleId = (string?)pPr?.Element(W.pStyle)?.Attribute(W.val);
            numPr = ResolveStyleNumPr(doc, styleId);
            fromStyle = numPr is not null;
        }

        if (numPr is null) return null;

        var numIdEl = numPr.Element(W.numId);
        if (numIdEl is null) return null;
        if (!int.TryParse((string?)numIdEl.Attribute(W.val), out var numId)) return null;

        var ilvlEl = numPr.Element(W.ilvl);
        int level = 0;
        if (ilvlEl is not null) int.TryParse((string?)ilvlEl.Attribute(W.val), out level);

        var numberingPart = doc.MainDocumentPart?.NumberingDefinitionsPart;
        var numberingRoot = numberingPart?.GetXDocument().Root;
        if (numberingRoot is null) return null;

        var numEl = numberingRoot.Elements(W.num)
            .FirstOrDefault(n => (string?)n.Attribute(W.numId) == numId.ToString(System.Globalization.CultureInfo.InvariantCulture));
        if (numEl is null) return null;

        var abstractNumIdEl = numEl.Element(W.abstractNumId);
        if (abstractNumIdEl is null) return null;
        if (!int.TryParse((string?)abstractNumIdEl.Attribute(W.val), out var abstractNumId)) return null;

        var abstractNumEl = numberingRoot.Elements(W.abstractNum)
            .FirstOrDefault(a => (string?)a.Attribute(W.abstractNumId) == abstractNumId.ToString(System.Globalization.CultureInfo.InvariantCulture));
        if (abstractNumEl is null) return null;

        var lvlEl = abstractNumEl.Elements(W.lvl)
            .FirstOrDefault(l => (string?)l.Attribute(W.ilvl) == level.ToString(System.Globalization.CultureInfo.InvariantCulture));
        var format = ParseNumberFormat((string?)lvlEl?.Element(W.numFmt)?.Attribute(W.val));

        // Start override from the w:num's lvlOverride for this level.
        int? startOverride = null;
        var lvlOverrideEl = numEl.Elements(W.lvlOverride)
            .FirstOrDefault(o => (string?)o.Attribute(W.ilvl) == level.ToString(System.Globalization.CultureInfo.InvariantCulture));
        if (lvlOverrideEl is not null)
        {
            var startOverrideEl = lvlOverrideEl.Element(W.startOverride);
            if (startOverrideEl is not null && int.TryParse((string?)startOverrideEl.Attribute(W.val), out var so))
                startOverride = so;
        }

        return new ListMembership
        {
            NumId = numId,
            AbstractNumId = abstractNumId,
            Level = level,
            Format = format,
            StartOverride = startOverride,
            IsAutoNumbered = true,
            FromStyle = fromStyle,
            GeneratedLabel = target.AutoNumberPrefix,
        };
    }

    private static XElement? ResolveStyleNumPr(WordprocessingDocument doc, string? styleId)
    {
        if (string.IsNullOrEmpty(styleId)) return null;
        var stylesPart = doc.MainDocumentPart?.StyleDefinitionsPart;
        var root = stylesPart?.GetXDocument().Root;
        if (root is null) return null;

        // Walk style → basedOn → ... up to 16 levels (cycle-guard depth).
        var visited = new HashSet<string>(System.StringComparer.Ordinal);
        var current = styleId;
        for (int i = 0; i < 16 && current is not null; i++)
        {
            if (!visited.Add(current)) return null;  // cycle
            var style = root.Elements(W.style).FirstOrDefault(s => (string?)s.Attribute(W.styleId) == current);
            if (style is null) return null;
            var numPr = style.Element(W.pPr)?.Element(W.numPr);
            if (numPr is not null) return numPr;
            current = (string?)style.Element(W.basedOn)?.Attribute(W.val);
        }
        return null;
    }

    private static NumberFormat ParseNumberFormat(string? raw) => raw switch
    {
        "bullet" => NumberFormat.Bullet,
        "upperLetter" => NumberFormat.UpperLetter,
        "lowerLetter" => NumberFormat.LowerLetter,
        "upperRoman" => NumberFormat.UpperRoman,
        "lowerRoman" => NumberFormat.LowerRoman,
        _ => NumberFormat.Decimal,  // includes "decimal" and any unrecognized format
    };
```

- [ ] **Step 4: Run the two tests to verify they pass**

```bash
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~BM002|FullyQualifiedName~BM003"
```
Expected: 2 passed.

- [ ] **Step 5: Commit**

```bash
git add Docxodus.Tests/DocxSessionMetadataTests.cs Docxodus/Internal/BlockMetadataOps.cs
git commit -m "feat(session): resolve ListMembership for inline numPr (BM002/BM003)"
```

---

### Task 6: Add tests + fixture for style-inherited `w:numPr` (`FromStyle = true`)

**Files:**
- Modify: `Docxodus.Tests/DocxSessionTests.cs` — add a new fixture builder
- Modify: `Docxodus.Tests/DocxSessionMetadataTests.cs` — add a test

- [ ] **Step 1: Add the new fixture builder in `DocxSessionTests.cs`**

Insert after `BuildDS002_BulletedList` (around line 120):

```csharp
    /// <summary>
    /// Single-item list where the paragraph carries only a <c>pStyle</c> pointing
    /// at a custom style that contributes the <c>numPr</c>. Used to verify
    /// <c>FromStyle = true</c> resolution.
    /// </summary>
    internal static byte[] BuildBM_StyleInheritedList()
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.Document = new Document();
            var body = new Body();
            main.Document.Body = body;

            var stylesPart = main.AddNewPart<StyleDefinitionsPart>();
            var styles = BuildHeadingStyles();
            // Add a paragraph style "MyListStyle" that carries numPr → numId=1, ilvl=0.
            var styleNumPr = new ParagraphProperties(
                new NumberingProperties(
                    new NumberingLevelReference { Val = 0 },
                    new NumberingId { Val = 1 }));
            styles.Append(new Style(
                new StyleName { Val = "My List Style" },
                styleNumPr)
            {
                Type = StyleValues.Paragraph,
                StyleId = "MyListStyle",
            });
            stylesPart.Styles = styles;

            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var numberingPart = main.AddNewPart<NumberingDefinitionsPart>();
            numberingPart.Numbering = BuildBulletNumbering();

            // Paragraph carries only the pStyle — no inline numPr.
            var pPr = new ParagraphProperties(new ParagraphStyleId { Val = "MyListStyle" });
            body.Append(new Paragraph(pPr, new Run(new Text("Style-inherited list item"))));

            main.Document.Save();
        }
        return ms.ToArray();
    }
```

- [ ] **Step 2: Add test BM004**

Append to `DocxSessionMetadataTests`:

```csharp
    [Fact]
    public void BM004_GetListMembership_StyleInheritedNumPr_SetsFromStyleTrue()
    {
        using var session = new DocxSession(DocxSessionTests.BuildBM_StyleInheritedList());
        var anchor = session.Project().AnchorIndex.Values.First(t => t.Anchor.Kind == "li");

        var list = session.GetListMembership(anchor.Anchor.Id);

        Assert.NotNull(list);
        Assert.True(list!.FromStyle);
        Assert.Equal(1, list.NumId);
        Assert.Equal(0, list.Level);
        Assert.Equal(NumberFormat.Bullet, list.Format);
    }
```

- [ ] **Step 3: Run BM004 to verify it passes**

```bash
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~BM004"
```
Expected: 1 passed.

- [ ] **Step 4: Commit**

```bash
git add Docxodus.Tests/DocxSessionTests.cs Docxodus.Tests/DocxSessionMetadataTests.cs
git commit -m "test(session): BM004 verifies style-inherited numPr sets FromStyle=true"
```

---

## Phase 3: Section info resolution

### Task 7: Implement `GetSectionInfo` for body anchors

**Files:**
- Modify: `Docxodus/Internal/BlockMetadataOps.cs`
- Modify: `Docxodus.Tests/DocxSessionTests.cs` — fixture with explicit `w:sectPr`
- Modify: `Docxodus.Tests/DocxSessionMetadataTests.cs` — tests

- [ ] **Step 1: Add fixture `BuildBM_LandscapeSection` in `DocxSessionTests.cs`** (insert after `BuildBM_StyleInheritedList`)

```csharp
    /// <summary>
    /// Single paragraph in a landscape A4 section with non-default margins,
    /// two columns, and a header part reference. Used to verify
    /// <c>GetSectionInfo</c> against a richly-configured sectPr.
    /// </summary>
    internal static byte[] BuildBM_LandscapeSection()
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.Document = new Document();
            var body = new Body();
            main.Document.Body = body;

            main.AddNewPart<StyleDefinitionsPart>().Styles = BuildHeadingStyles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            // Add a header part so we can verify HeaderPartUris is populated.
            var headerPart = main.AddNewPart<HeaderPart>("rIdH1");
            headerPart.Header = new Header(new Paragraph(new Run(new Text("Header text"))));

            body.Append(new Paragraph(new Run(new Text("Page body"))));

            // Section properties: A4 landscape (16838 x 11906 twips), columns=2, margins set.
            var sectPr = new SectionProperties(
                new HeaderReference { Type = HeaderFooterValues.Default, Id = "rIdH1" },
                new PageSize { Width = 16838u, Height = 11906u, Orient = PageOrientationValues.Landscape },
                new PageMargin { Top = 720, Bottom = 720, Left = 1080, Right = 1080,
                    Header = 0, Footer = 0, Gutter = 0 },
                new Columns { ColumnCount = 2 });
            body.Append(sectPr);

            main.Document.Save();
        }
        return ms.ToArray();
    }
```

- [ ] **Step 2: Add failing tests BM005/BM006**

```csharp
    [Fact]
    public void BM005_GetSectionInfo_BodyAnchor_ResolvesLandscapeAndHeaders()
    {
        using var session = new DocxSession(DocxSessionTests.BuildBM_LandscapeSection());
        var anchor = session.Project().AnchorIndex.Values.First(t => t.Anchor.Kind == "p");

        var info = session.GetSectionInfo(anchor.Anchor.Id);

        Assert.NotNull(info);
        Assert.Equal(16838.0, info!.PageWidthTwips);
        Assert.Equal(11906.0, info.PageHeightTwips);
        Assert.True(info.Landscape);
        Assert.Equal(720.0, info.MarginTopTwips);
        Assert.Equal(720.0, info.MarginBottomTwips);
        Assert.Equal(1080.0, info.MarginLeftTwips);
        Assert.Equal(1080.0, info.MarginRightTwips);
        Assert.Equal(2, info.Columns);
        Assert.Single(info.HeaderPartUris);
        Assert.Empty(info.FooterPartUris);
    }

    [Fact]
    public void BM006_GetSectionInfo_UnknownAnchor_ReturnsNull()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        Assert.Null(session.GetSectionInfo("p:body:does-not-exist"));
    }
```

- [ ] **Step 3: Run to verify BM005 fails** (the resolver is still the Task-3 stub returning null)

```bash
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~BM005|FullyQualifiedName~BM006"
```
Expected: BM006 PASS, BM005 FAIL (`Assert.NotNull(info)` fails).

- [ ] **Step 4: Implement `GetSectionInfo` in `BlockMetadataOps.cs`**

Replace the stub `GetSectionInfo` with:

```csharp
    public static SectionInfo? GetSectionInfo(WordprocessingDocument doc, AnchorTarget target)
    {
        // Only body anchors have a sectPr to look up.
        if (target.Anchor.Scope != "body") return null;

        var element = target.Resolve(doc);
        if (element is null) return null;

        var sectPr = FindGoverningSectPr(element);
        if (sectPr is null) return null;

        var pgSz = sectPr.Element(W.pgSz);
        var pgMar = sectPr.Element(W.pgMar);
        var cols = sectPr.Element(W.cols);

        // The width attribute is `w:w` — exposed in the W class as `W._w` to avoid
        // collision with the namespace alias `W.w`. Height is just `W.h`. See
        // WmlToHtmlConverter for the same usage pattern.
        double width = ParseDouble((string?)pgSz?.Attribute(W._w)) ?? 12240;   // 8.5"
        double height = ParseDouble((string?)pgSz?.Attribute(W.h)) ?? 15840;  // 11"
        bool landscape = string.Equals((string?)pgSz?.Attribute(W.orient), "landscape",
            System.StringComparison.Ordinal);

        double top = ParseDouble((string?)pgMar?.Attribute(W.top)) ?? 1440;
        double bottom = ParseDouble((string?)pgMar?.Attribute(W.bottom)) ?? 1440;
        double left = ParseDouble((string?)pgMar?.Attribute(W.left)) ?? 1440;
        double right = ParseDouble((string?)pgMar?.Attribute(W.right)) ?? 1440;

        int colCount = 1;
        if (cols is not null && int.TryParse((string?)cols.Attribute(W.num), out var parsedCols))
            colCount = parsedCols;

        var (headerUris, footerUris) = ResolveSectionHeaderFooterUris(doc, sectPr);

        // The sectPr itself doesn't carry a stable Unid in every fixture; fall back
        // to a deterministic synthetic id derived from element position so the field
        // is always non-null and stable across reads of the same doc state.
        var sectionUnid = (string?)sectPr.Attribute(PtOpenXml.Unid)
            ?? $"sect:{sectPr.Parent?.Elements().ToList().IndexOf(sectPr) ?? 0}";

        return new SectionInfo
        {
            SectionUnid = sectionUnid,
            PageWidthTwips = width,
            PageHeightTwips = height,
            Landscape = landscape,
            MarginTopTwips = top,
            MarginBottomTwips = bottom,
            MarginLeftTwips = left,
            MarginRightTwips = right,
            Columns = colCount,
            HeaderPartUris = headerUris,
            FooterPartUris = footerUris,
        };
    }

    private static XElement? FindGoverningSectPr(XElement element)
    {
        // The sectPr that governs a paragraph is either:
        // (a) the next sectPr-bearing paragraph's pPr/sectPr after this one (mid-doc section break), or
        // (b) the body's trailing sectPr (the document-final section).
        var body = element.AncestorsAndSelf(W.body).FirstOrDefault();
        if (body is null) return null;

        // Walk forward from the element looking for any p whose pPr has a sectPr.
        foreach (var sib in element.ElementsAfterSelf())
        {
            if (sib.Name != W.p) continue;
            var sp = sib.Element(W.pPr)?.Element(W.sectPr);
            if (sp is not null) return sp;
        }
        if (element.Name == W.p)
        {
            var ownSect = element.Element(W.pPr)?.Element(W.sectPr);
            if (ownSect is not null) return ownSect;
        }
        return body.Element(W.sectPr);
    }

    private static (IReadOnlyList<string> headers, IReadOnlyList<string> footers)
        ResolveSectionHeaderFooterUris(WordprocessingDocument doc, XElement sectPr)
    {
        var main = doc.MainDocumentPart;
        if (main is null)
            return (System.Array.Empty<string>(), System.Array.Empty<string>());

        var headers = new List<string>();
        var footers = new List<string>();

        foreach (var headerRef in sectPr.Elements(W.headerReference))
        {
            var rId = (string?)headerRef.Attribute(R.id);
            if (rId is null) continue;
            var part = main.GetPartById(rId) as HeaderPart;
            if (part is not null) headers.Add(part.Uri.ToString());
        }
        foreach (var footerRef in sectPr.Elements(W.footerReference))
        {
            var rId = (string?)footerRef.Attribute(R.id);
            if (rId is null) continue;
            var part = main.GetPartById(rId) as FooterPart;
            if (part is not null) footers.Add(part.Uri.ToString());
        }
        return (headers, footers);
    }

    private static double? ParseDouble(string? raw)
        => double.TryParse(raw, System.Globalization.NumberStyles.Any,
            System.Globalization.CultureInfo.InvariantCulture, out var v) ? v : null;
```

- [ ] **Step 5: Verify the build compiles**

All XName references used in this resolver (`W._w`, `W.h`, `W.top`, `W.bottom`, `W.left`, `W.right`, `W.num`, `W.headerReference`, `W.footerReference`, `R.id`, `W.lvlOverride`, `W.startOverride`) already exist in `Docxodus/PtOpenXmlUtil.cs` (verified before plan-writing).

```bash
dotnet build Docxodus/Docxodus.csproj 2>&1 | grep -E "error|Build succeeded" | head -10
```
Expected: "Build succeeded".

- [ ] **Step 6: Run the failing tests to verify they pass**

```bash
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~BM005|FullyQualifiedName~BM006"
```
Expected: 2 passed.

- [ ] **Step 7: Commit**

```bash
git add Docxodus/Internal/BlockMetadataOps.cs Docxodus/PtOpenXmlUtil.cs Docxodus.Tests/DocxSessionTests.cs Docxodus.Tests/DocxSessionMetadataTests.cs
git commit -m "feat(session): resolve SectionInfo (page size/orient/margins/cols/headers/footers)"
```

---

### Task 8: Add tests for `SectionInfo` returning null on non-body anchors

**Files:**
- Modify: `Docxodus.Tests/DocxSessionMetadataTests.cs`

The implementation already returns null for non-body scopes; this just locks it down with a test.

- [ ] **Step 1: Add test BM007**

```csharp
    [Fact]
    public void BM007_GetSectionInfo_NonBodyAnchor_ReturnsNull()
    {
        // The landscape-section fixture has a HeaderPart with one paragraph.
        // That paragraph's anchor lives in scope "hdr1", not "body".
        using var session = new DocxSession(DocxSessionTests.BuildBM_LandscapeSection());
        var hdrAnchor = session.Project().AnchorIndex.Values
            .FirstOrDefault(t => t.Anchor.Scope.StartsWith("hdr", System.StringComparison.Ordinal));
        Assert.NotNull(hdrAnchor);

        Assert.Null(session.GetSectionInfo(hdrAnchor!.Anchor.Id));
    }
```

- [ ] **Step 2: Run to verify it passes**

```bash
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~BM007"
```
Expected: 1 passed.

- [ ] **Step 3: Commit**

```bash
git add Docxodus.Tests/DocxSessionMetadataTests.cs
git commit -m "test(session): BM007 verifies GetSectionInfo returns null for non-body anchors"
```

---

## Phase 4: Bulk reads, outline level, formatting probe

### Task 9: Add tests + verify outline level inference and bulk reads

**Files:**
- Modify: `Docxodus.Tests/DocxSessionMetadataTests.cs`

The implementation in Task 3/4 already covers outline level + bulk reads; these tests lock the behavior down.

- [ ] **Step 1: Add tests BM008/BM009/BM010**

```csharp
    [Fact]
    public void BM008_OutlineLevel_FromHeadingStyle_ResolvesToZeroBasedLevel()
    {
        // BuildDS001 has Heading1..6 styles defined. Apply Heading2 to the first paragraph.
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = session.Project().AnchorIndex.Values.First(t => t.Anchor.Kind == "p");
        var setStyle = session.SetParagraphStyle(anchor.Anchor.Id, "Heading2");
        Assert.True(setStyle.Success);

        // SetParagraphStyle may have changed the anchor kind from "p" to "h" — re-resolve.
        var freshIndex = session.Project().AnchorIndex;
        var promoted = freshIndex.Values.First(t => t.Anchor.Kind == "h");

        var meta = session.GetBlockMetadata(promoted.Anchor.Id);
        Assert.NotNull(meta);
        Assert.Equal("Heading2", meta!.StyleId);
        Assert.Equal("Heading 2", meta.StyleName);
        Assert.Equal(1, meta.OutlineLevel);  // Heading2 → outlineLvl 1 (0-based)
    }

    [Fact]
    public void BM009_GetBlockMetadatas_Bulk_DedupesAndMapsUnknownToNull()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchors = session.Project().AnchorIndex.Values.Where(t => t.Anchor.Kind == "p").ToList();
        Assert.True(anchors.Count >= 2);

        var ids = new[] {
            anchors[0].Anchor.Id,
            anchors[0].Anchor.Id,         // duplicate
            anchors[1].Anchor.Id,
            "p:body:does-not-exist",
        };

        var map = session.GetBlockMetadatas(ids);

        Assert.Equal(3, map.Count);  // duplicate dropped
        Assert.NotNull(map[anchors[0].Anchor.Id]);
        Assert.NotNull(map[anchors[1].Anchor.Id]);
        Assert.Null(map["p:body:does-not-exist"]);
    }

    [Fact]
    public void BM010_HasInlineFormatting_DetectsBoldRun()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = session.Project().AnchorIndex.Values.First(t => t.Anchor.Kind == "p");

        Assert.False(session.GetBlockMetadata(anchor.Anchor.Id)!.HasInlineFormatting);

        var apply = session.ApplyFormat(anchor.Anchor.Id, span: null, new FormatOp { Bold = true });
        Assert.True(apply.Success);

        Assert.True(session.GetBlockMetadata(anchor.Anchor.Id)!.HasInlineFormatting);
    }
```

- [ ] **Step 2: Run all BM tests**

```bash
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~BM00|FullyQualifiedName~BM010"
```
Expected: 10 passed (BM001-BM010).

- [ ] **Step 3: Commit**

```bash
git add Docxodus.Tests/DocxSessionMetadataTests.cs
git commit -m "test(session): BM008/9/10 cover outline level, bulk read dedup, formatting probe"
```

---

## Phase 5: Wire through DocxSessionOps + JSON facade

### Task 10: Add `DocxSessionOps` facade methods + JSON serializers

**Files:**
- Modify: `Docxodus/Internal/DocxSessionOps.cs`
- Modify: `Docxodus/Internal/DocxSessionJson.cs`

- [ ] **Step 1: Add the four facade methods to `DocxSessionOps.cs`**

Insert after the existing `GetAnchorInfos` facade method (around line 67) and before `FindByText`:

```csharp
    public static string GetBlockMetadata(int handle, string anchorId) =>
        DocxSessionJson.SerializeBlockMetadataOrNull(SessionRegistry.Get(handle).GetBlockMetadata(anchorId));

    public static string GetBlockMetadatas(int handle, System.Collections.Generic.IEnumerable<string> anchorIds) =>
        DocxSessionJson.SerializeBlockMetadataMap(SessionRegistry.Get(handle).GetBlockMetadatas(anchorIds));

    public static string GetListMembership(int handle, string anchorId) =>
        DocxSessionJson.SerializeListMembershipOrNull(SessionRegistry.Get(handle).GetListMembership(anchorId));

    public static string GetSectionInfo(int handle, string anchorId) =>
        DocxSessionJson.SerializeSectionInfoOrNull(SessionRegistry.Get(handle).GetSectionInfo(anchorId));
```

- [ ] **Step 2: Add the serializers to `DocxSessionJson.cs`**

Insert in the "Serializers" region (after `SerializeAnchorInfoOrNull` around line 524). Reuse the existing `JsonString` helper for string escaping.

```csharp
    public static string SerializeBlockMetadataOrNull(BlockMetadata? meta)
    {
        if (meta is null) return "null";
        var sb = new StringBuilder(256);
        sb.Append("{\"anchorId\":").Append(JsonString(meta.AnchorId))
          .Append(",\"kind\":").Append(JsonString(meta.Kind))
          .Append(",\"scope\":").Append(JsonString(meta.Scope));
        if (meta.StyleId is not null)
            sb.Append(",\"styleId\":").Append(JsonString(meta.StyleId));
        if (meta.StyleName is not null)
            sb.Append(",\"styleName\":").Append(JsonString(meta.StyleName));
        if (meta.OutlineLevel.HasValue)
            sb.Append(",\"outlineLevel\":").Append(meta.OutlineLevel.Value);
        if (meta.List is not null)
            sb.Append(",\"list\":").Append(SerializeListMembershipOrNull(meta.List));
        sb.Append(",\"hasInlineFormatting\":").Append(meta.HasInlineFormatting ? "true" : "false");
        sb.Append('}');
        return sb.ToString();
    }

    public static string SerializeBlockMetadataMap(System.Collections.Generic.IReadOnlyDictionary<string, BlockMetadata?> map)
    {
        var sb = new StringBuilder(map.Count * 200 + 2);
        sb.Append('{');
        bool first = true;
        foreach (var kv in map)
        {
            if (!first) sb.Append(',');
            first = false;
            sb.Append(JsonString(kv.Key)).Append(':');
            sb.Append(SerializeBlockMetadataOrNull(kv.Value));
        }
        sb.Append('}');
        return sb.ToString();
    }

    public static string SerializeListMembershipOrNull(ListMembership? list)
    {
        if (list is null) return "null";
        var sb = new StringBuilder(128);
        sb.Append("{\"numId\":").Append(list.NumId)
          .Append(",\"abstractNumId\":").Append(list.AbstractNumId)
          .Append(",\"level\":").Append(list.Level)
          .Append(",\"format\":").Append(JsonString(NumberFormatToString(list.Format)))
          .Append(",\"isAutoNumbered\":").Append(list.IsAutoNumbered ? "true" : "false")
          .Append(",\"fromStyle\":").Append(list.FromStyle ? "true" : "false");
        if (list.StartOverride.HasValue)
            sb.Append(",\"startOverride\":").Append(list.StartOverride.Value);
        if (list.GeneratedLabel is not null)
            sb.Append(",\"generatedLabel\":").Append(JsonString(list.GeneratedLabel));
        sb.Append('}');
        return sb.ToString();
    }

    public static string SerializeSectionInfoOrNull(SectionInfo? info)
    {
        if (info is null) return "null";
        var sb = new StringBuilder(256);
        sb.Append("{\"sectionUnid\":").Append(JsonString(info.SectionUnid))
          .Append(",\"pageWidthTwips\":").Append(info.PageWidthTwips.ToString(System.Globalization.CultureInfo.InvariantCulture))
          .Append(",\"pageHeightTwips\":").Append(info.PageHeightTwips.ToString(System.Globalization.CultureInfo.InvariantCulture))
          .Append(",\"landscape\":").Append(info.Landscape ? "true" : "false")
          .Append(",\"marginTopTwips\":").Append(info.MarginTopTwips.ToString(System.Globalization.CultureInfo.InvariantCulture))
          .Append(",\"marginBottomTwips\":").Append(info.MarginBottomTwips.ToString(System.Globalization.CultureInfo.InvariantCulture))
          .Append(",\"marginLeftTwips\":").Append(info.MarginLeftTwips.ToString(System.Globalization.CultureInfo.InvariantCulture))
          .Append(",\"marginRightTwips\":").Append(info.MarginRightTwips.ToString(System.Globalization.CultureInfo.InvariantCulture))
          .Append(",\"columns\":").Append(info.Columns)
          .Append(",\"headerPartUris\":[");
        for (int i = 0; i < info.HeaderPartUris.Count; i++)
        {
            if (i > 0) sb.Append(',');
            sb.Append(JsonString(info.HeaderPartUris[i]));
        }
        sb.Append("],\"footerPartUris\":[");
        for (int i = 0; i < info.FooterPartUris.Count; i++)
        {
            if (i > 0) sb.Append(',');
            sb.Append(JsonString(info.FooterPartUris[i]));
        }
        sb.Append("]}");
        return sb.ToString();
    }

    private static string NumberFormatToString(NumberFormat f) => f switch
    {
        NumberFormat.Decimal => "decimal",
        NumberFormat.UpperLetter => "upperLetter",
        NumberFormat.LowerLetter => "lowerLetter",
        NumberFormat.UpperRoman => "upperRoman",
        NumberFormat.LowerRoman => "lowerRoman",
        NumberFormat.Bullet => "bullet",
        _ => "decimal",
    };
```

- [ ] **Step 3: Verify the build**

```bash
dotnet build Docxodus/Docxodus.csproj
```
Expected: build succeeds.

- [ ] **Step 4: Commit**

```bash
git add Docxodus/Internal/DocxSessionOps.cs Docxodus/Internal/DocxSessionJson.cs
git commit -m "feat(ops): expose block-metadata reads + JSON serializers"
```

---

## Phase 6: WASM bridge + stdio dispatcher

### Task 11: Add WASM `[JSExport]` shells

**Files:**
- Modify: `wasm/DocxodusWasm/DocxSessionBridge.cs`

- [ ] **Step 1: Locate the insertion point**

```bash
grep -n "public static string GetAnchorInfo\|public static string GetAnchorInfos" wasm/DocxodusWasm/DocxSessionBridge.cs
```
Expected: hits around lines 320 and 328.

- [ ] **Step 2: Insert the four shells after the existing `GetAnchorInfos` shell**

```csharp
    /// <summary>
    /// Bridge for <see cref="DocxSession.GetBlockMetadata"/>. Returns a JSON
    /// object with style id/name, outline level, list membership (when present),
    /// and a hasInlineFormatting flag — or <c>"null"</c> if the anchor doesn't exist.
    /// </summary>
    [JSExport]
    public static string GetBlockMetadata(int h, string anchorId) =>
        DocxSessionOps.GetBlockMetadata(h, anchorId);

    /// <summary>
    /// Bulk variant of <see cref="GetBlockMetadata"/>. Takes a JSON array of anchor
    /// ids, returns a JSON object mapping each id to its metadata (or null).
    /// </summary>
    [JSExport]
    public static string GetBlockMetadatas(int h, string anchorIdsJson)
    {
        var ids = System.Text.Json.JsonSerializer.Deserialize<string[]>(anchorIdsJson)
            ?? System.Array.Empty<string>();
        return DocxSessionOps.GetBlockMetadatas(h, ids);
    }

    /// <summary>
    /// Bridge for <see cref="DocxSession.GetListMembership"/>. Returns a JSON
    /// object with numId/abstractNumId/level/format/etc., or <c>"null"</c>.
    /// </summary>
    [JSExport]
    public static string GetListMembership(int h, string anchorId) =>
        DocxSessionOps.GetListMembership(h, anchorId);

    /// <summary>
    /// Bridge for <see cref="DocxSession.GetSectionInfo"/>. Returns a JSON object
    /// describing the governing <c>w:sectPr</c>, or <c>"null"</c> for non-body anchors.
    /// </summary>
    [JSExport]
    public static string GetSectionInfo(int h, string anchorId) =>
        DocxSessionOps.GetSectionInfo(h, anchorId);
```

- [ ] **Step 3: Build the WASM target to verify**

```bash
./scripts/build-wasm.sh 2>&1 | tail -5
```
Expected: build succeeds.

- [ ] **Step 4: Restore non-WASM build state for further .NET work**

```bash
dotnet clean Docxodus.sln && dotnet build Docxodus/Docxodus.csproj
```
Expected: clean + rebuild succeed (this avoids the WASM/non-WASM cached-DLL mismatch documented in CLAUDE.md).

- [ ] **Step 5: Commit**

```bash
git add wasm/DocxodusWasm/DocxSessionBridge.cs
git commit -m "feat(wasm): expose block-metadata reads via JSExport"
```

---

### Task 12: Add stdio host dispatcher cases

**Files:**
- Modify: `tools/python-host/Dispatcher.cs`

- [ ] **Step 1: Locate the insertion point**

```bash
grep -n "\"get_anchor_info\"\|\"get_anchor_infos\"" tools/python-host/Dispatcher.cs
```
Expected: hits around lines 111-112.

- [ ] **Step 2: Insert the four new cases after `"get_anchor_infos"`**

```csharp
        "get_block_metadata" => DocxSessionOps.GetBlockMetadata(Handle(args), Str(args, "anchorId")),
        "get_block_metadatas" => DocxSessionOps.GetBlockMetadatas(Handle(args), ParseAnchorIdArray(args)),
        "get_list_membership" => DocxSessionOps.GetListMembership(Handle(args), Str(args, "anchorId")),
        "get_section_info" => DocxSessionOps.GetSectionInfo(Handle(args), Str(args, "anchorId")),
```

- [ ] **Step 3: Build the python-host binary to verify**

```bash
dotnet build tools/python-host/python-host.csproj
```
Expected: build succeeds.

- [ ] **Step 4: Commit**

```bash
git add tools/python-host/Dispatcher.cs
git commit -m "feat(pyhost): dispatch block-metadata read ops"
```

---

## Phase 7: npm wrapper

### Task 13: Add TypeScript types

**Files:**
- Modify: `npm/src/types.ts`
- Modify: `npm/src/index.ts`

- [ ] **Step 1: Insert the type definitions** — find the existing `AnchorInfo` interface block (around line 1189) and insert the new types immediately after it:

```typescript
/** Six list formats supported by the list write surface (decimal, upperLetter,
 *  lowerLetter, upperRoman, lowerRoman, bullet). Surfaced on
 *  {@link ListMembership.format} as a string union (mirrors the JSON wire format). */
export type NumberFormat =
  | "decimal"
  | "upperLetter"
  | "lowerLetter"
  | "upperRoman"
  | "lowerRoman"
  | "bullet";

/** Numbering facts for a list-item paragraph. Returned by
 *  {@link DocxSession.getListMembership} and surfaced as {@link BlockMetadata.list}. */
export interface ListMembership {
  /** The w:numId the paragraph belongs to (the w:num instance). */
  numId: number;
  /** The w:abstractNumId the paragraph's w:num points at. */
  abstractNumId: number;
  /** The paragraph's level (w:ilvl), 0-8. */
  level: number;
  /** Resolved format for this level. */
  format: NumberFormat;
  /** Always true for a paragraph carrying w:numPr (inline or via style). */
  isAutoNumbered: boolean;
  /** True when the w:numPr is inherited from the paragraph's style chain. */
  fromStyle: boolean;
  /** Start-override from w:lvlOverride/w:startOverride for this level, if any. */
  startOverride?: number;
  /** Resolved label (e.g. "1.", "(a)") — same value surfaced via AnchorInfo.autoNumberPrefix. */
  generatedLabel?: string;
}

/** Block-level structural metadata. Returned by {@link DocxSession.getBlockMetadata}. */
export interface BlockMetadata {
  anchorId: string;
  kind: string;
  scope: string;
  styleId?: string;
  styleName?: string;
  /** 0-based outline level (Word convention). */
  outlineLevel?: number;
  list?: ListMembership;
  /** True when any descendant w:r carries a non-empty w:rPr. */
  hasInlineFormatting: boolean;
}

/** Page-layout snapshot for the w:sectPr that governs an anchor.
 *  Returned by {@link DocxSession.getSectionInfo}. */
export interface SectionInfo {
  sectionUnid: string;
  pageWidthTwips: number;
  pageHeightTwips: number;
  landscape: boolean;
  marginTopTwips: number;
  marginBottomTwips: number;
  marginLeftTwips: number;
  marginRightTwips: number;
  columns: number;
  headerPartUris: string[];
  footerPartUris: string[];
}
```

- [ ] **Step 2: Add the four wire signatures to `DocxodusWasmExports`** — find the existing `GetAnchorInfos:` line (around line 767) and insert after it:

```typescript
    GetBlockMetadata: (handle: number, anchorId: string) => string;
    GetBlockMetadatas: (handle: number, anchorIdsJson: string) => string;
    GetListMembership: (handle: number, anchorId: string) => string;
    GetSectionInfo: (handle: number, anchorId: string) => string;
```

- [ ] **Step 3: Re-export the new types from `npm/src/index.ts`** — find the existing `export type { ... AnchorInfo ... } from "./types.js"` block (around line 56-68) and add the new types to the export list:

```typescript
export type {
  AnchorInfo,
  AnchorRef,
  AnchorTargetRef,
  BlockMetadata,
  CharSpan,
  DocumentAnnotation,
  DocxSessionProjection,
  DocxSessionSettings,
  EditError,
  EditErrorCode,
  EditResult,
  FormatOp,
  ListMembership,
  MarkdownPatch,
  NumberFormat,
  SectionInfo,
} from "./types.js";
```

- [ ] **Step 4: Type-check**

```bash
cd npm && npx tsc --noEmit
```
Expected: 0 errors.

- [ ] **Step 5: Commit**

```bash
git add npm/src/types.ts npm/src/index.ts
git commit -m "feat(npm): add BlockMetadata / ListMembership / SectionInfo TypeScript types"
```

---

### Task 14: Add wrapper methods on the npm `DocxSession` class

**Files:**
- Modify: `npm/src/session.ts`

- [ ] **Step 1: Add the four methods after the existing `getAnchorInfos` method (around line 544)**

```typescript
  /**
   * Resolve block-level metadata (style id+name, outline level, list membership,
   * formatting probe) for an anchor. Returns null when the anchor doesn't exist.
   */
  getBlockMetadata(anchorId: string): BlockMetadata | null {
    const raw = this.wasm.GetBlockMetadata(this.handle, anchorId);
    return JSON.parse(raw) as BlockMetadata | null;
  }

  /**
   * Bulk variant of {@link getBlockMetadata}. Unknown ids map to null;
   * duplicates are deduped.
   */
  getBlockMetadatas(anchorIds: readonly string[]): Record<string, BlockMetadata | null> {
    const raw = this.wasm.GetBlockMetadatas(this.handle, JSON.stringify(anchorIds));
    return JSON.parse(raw) as Record<string, BlockMetadata | null>;
  }

  /**
   * Resolve the numbering facts for a list-item paragraph; returns null when
   * the anchor has no w:numPr.
   */
  getListMembership(anchorId: string): ListMembership | null {
    const raw = this.wasm.GetListMembership(this.handle, anchorId);
    return JSON.parse(raw) as ListMembership | null;
  }

  /**
   * Resolve page-layout info for the w:sectPr that governs an anchor.
   * Returns null for anchors outside the body part.
   */
  getSectionInfo(anchorId: string): SectionInfo | null {
    const raw = this.wasm.GetSectionInfo(this.handle, anchorId);
    return JSON.parse(raw) as SectionInfo | null;
  }
```

- [ ] **Step 2: Add the four new types to the import line at the top of `session.ts`**

Find the existing import block and add `BlockMetadata`, `ListMembership`, `SectionInfo` to it (alphabetized).

```bash
grep -n "^import type \{" npm/src/session.ts | head -3
```

Edit the import to include the new types.

- [ ] **Step 3: Type-check**

```bash
cd npm && npx tsc --noEmit
```
Expected: 0 errors.

- [ ] **Step 4: Commit**

```bash
git add npm/src/session.ts
git commit -m "feat(npm): expose getBlockMetadata / getListMembership / getSectionInfo"
```

---

### Task 15: Add Playwright spec

**Files:**
- Create: `npm/tests/block-metadata.spec.ts`

- [ ] **Step 1: Build the npm package so Playwright has dist/ to load**

```bash
cd npm && npm run build 2>&1 | tail -5
```
Expected: build succeeds.

- [ ] **Step 2: Create the spec file**

```typescript
import { test, expect } from "@playwright/test";
import { readFileSync } from "node:fs";
import { resolve } from "node:path";

import { initialize, openDocxSession } from "../dist/index.js";

const fixtureDocx = readFileSync(
  resolve(__dirname, "fixtures/list-and-headings.docx")
);

test.beforeAll(async () => {
  await initialize();
});

test("getBlockMetadata returns kind+scope for plain paragraph", async () => {
  const session = await openDocxSession(fixtureDocx);
  try {
    const projection = session.project();
    const para = Object.values(projection.anchorIndex).find(
      (a) => a.kind === "p"
    )!;
    const meta = session.getBlockMetadata(para.id);
    expect(meta).not.toBeNull();
    expect(meta!.kind).toBe("p");
    expect(meta!.scope).toBe("body");
  } finally {
    session.close();
  }
});

test("getBlockMetadata returns null for unknown anchor", async () => {
  const session = await openDocxSession(fixtureDocx);
  try {
    expect(session.getBlockMetadata("p:body:does-not-exist")).toBeNull();
  } finally {
    session.close();
  }
});

test("getBlockMetadatas dedupes and maps unknown ids to null", async () => {
  const session = await openDocxSession(fixtureDocx);
  try {
    const projection = session.project();
    const paras = Object.values(projection.anchorIndex).filter(
      (a) => a.kind === "p"
    );
    expect(paras.length).toBeGreaterThanOrEqual(1);

    const map = session.getBlockMetadatas([
      paras[0].id,
      paras[0].id,
      "p:body:unknown",
    ]);
    expect(Object.keys(map).length).toBe(2);
    expect(map[paras[0].id]).not.toBeNull();
    expect(map["p:body:unknown"]).toBeNull();
  } finally {
    session.close();
  }
});

test("getListMembership returns null for non-list paragraphs", async () => {
  const session = await openDocxSession(fixtureDocx);
  try {
    const projection = session.project();
    const para = Object.values(projection.anchorIndex).find(
      (a) => a.kind === "p"
    )!;
    expect(session.getListMembership(para.id)).toBeNull();
  } finally {
    session.close();
  }
});

test("getSectionInfo returns page info for body anchors", async () => {
  const session = await openDocxSession(fixtureDocx);
  try {
    const projection = session.project();
    const para = Object.values(projection.anchorIndex).find(
      (a) => a.kind === "p"
    )!;
    const info = session.getSectionInfo(para.id);
    expect(info).not.toBeNull();
    expect(info!.pageWidthTwips).toBeGreaterThan(0);
    expect(info!.pageHeightTwips).toBeGreaterThan(0);
    expect(info!.columns).toBeGreaterThanOrEqual(1);
  } finally {
    session.close();
  }
});
```

- [ ] **Step 3: Verify a suitable fixture exists, or create one**

```bash
ls npm/tests/fixtures/ 2>/dev/null
```

If `list-and-headings.docx` isn't present, copy an existing fixture from `TestFiles/` that has both a list and at least one paragraph:

```bash
mkdir -p npm/tests/fixtures
cp TestFiles/DB012-Lists-With-Different-Numberings.docx npm/tests/fixtures/list-and-headings.docx
```

- [ ] **Step 4: Run the Playwright spec**

```bash
cd npm && npx playwright test block-metadata
```
Expected: 5 passed.

- [ ] **Step 5: Commit**

```bash
git add npm/tests/block-metadata.spec.ts npm/tests/fixtures/list-and-headings.docx
git commit -m "test(npm): Playwright coverage for block-metadata reads"
```

---

## Phase 8: Python wrapper

### Task 16: Add Python types

**Files:**
- Modify: `python/src/docx_scalpel/types.py`

- [ ] **Step 1: Find the insertion point**

```bash
grep -n "^class AnchorInfo:" python/src/docx_scalpel/types.py
```
Expected: one hit around line 110.

- [ ] **Step 2: Add the `Enum` import to the top of `types.py`**

The existing file imports only `from dataclasses import dataclass, field` and `from typing import Any, Mapping, Sequence` (verified). The new `NumberFormat` enum requires `Enum`. Add this line in the imports block (alphabetized with the others):

```python
from enum import Enum
```

- [ ] **Step 3: Add the new dataclasses + enum** — insert after the `AnchorInfo` class block (around line 125).

```python
class NumberFormat(str, Enum):
    """Six list formats supported by the list write surface and surfaced
    on ``ListMembership.format``. String-valued so the wire JSON round-trips
    transparently."""

    DECIMAL = "decimal"
    UPPER_LETTER = "upperLetter"
    LOWER_LETTER = "lowerLetter"
    UPPER_ROMAN = "upperRoman"
    LOWER_ROMAN = "lowerRoman"
    BULLET = "bullet"

    @classmethod
    def _from_wire(cls, raw: str) -> "NumberFormat":
        try:
            return cls(raw)
        except ValueError:
            return cls.DECIMAL


@dataclass(frozen=True, slots=True)
class ListMembership:
    """Numbering facts for a list-item paragraph."""

    num_id: int
    abstract_num_id: int
    level: int
    format: NumberFormat
    is_auto_numbered: bool
    from_style: bool
    start_override: int | None = None
    generated_label: str | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "ListMembership":
        return cls(
            num_id=int(d["numId"]),
            abstract_num_id=int(d["abstractNumId"]),
            level=int(d["level"]),
            format=NumberFormat._from_wire(d["format"]),
            is_auto_numbered=bool(d["isAutoNumbered"]),
            from_style=bool(d["fromStyle"]),
            start_override=int(d["startOverride"]) if "startOverride" in d else None,
            generated_label=d.get("generatedLabel"),
        )


@dataclass(frozen=True, slots=True)
class BlockMetadata:
    """Block-level structural metadata."""

    anchor_id: str
    kind: str
    scope: str
    has_inline_formatting: bool
    style_id: str | None = None
    style_name: str | None = None
    outline_level: int | None = None
    list: ListMembership | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "BlockMetadata":
        return cls(
            anchor_id=d["anchorId"],
            kind=d["kind"],
            scope=d["scope"],
            has_inline_formatting=bool(d["hasInlineFormatting"]),
            style_id=d.get("styleId"),
            style_name=d.get("styleName"),
            outline_level=int(d["outlineLevel"]) if "outlineLevel" in d else None,
            list=ListMembership._from_wire(d["list"]) if "list" in d else None,
        )


@dataclass(frozen=True, slots=True)
class SectionInfo:
    """Page-layout snapshot for the w:sectPr that governs an anchor."""

    section_unid: str
    page_width_twips: float
    page_height_twips: float
    landscape: bool
    margin_top_twips: float
    margin_bottom_twips: float
    margin_left_twips: float
    margin_right_twips: float
    columns: int
    header_part_uris: tuple[str, ...]
    footer_part_uris: tuple[str, ...]

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "SectionInfo":
        return cls(
            section_unid=d["sectionUnid"],
            page_width_twips=float(d["pageWidthTwips"]),
            page_height_twips=float(d["pageHeightTwips"]),
            landscape=bool(d["landscape"]),
            margin_top_twips=float(d["marginTopTwips"]),
            margin_bottom_twips=float(d["marginBottomTwips"]),
            margin_left_twips=float(d["marginLeftTwips"]),
            margin_right_twips=float(d["marginRightTwips"]),
            columns=int(d["columns"]),
            header_part_uris=tuple(d["headerPartUris"]),
            footer_part_uris=tuple(d["footerPartUris"]),
        )
```

- [ ] **Step 4: Add the four types to `__all__`** — find the existing `__all__` list at the top of the file and add `"BlockMetadata"`, `"ListMembership"`, `"SectionInfo"`, `"NumberFormat"` (alphabetized).

- [ ] **Step 5: Verify the module imports cleanly**

```bash
cd python && python -c "from docx_scalpel.types import BlockMetadata, ListMembership, SectionInfo, NumberFormat; print('ok')"
```
Expected: `ok`.

- [ ] **Step 6: Commit**

```bash
git add python/src/docx_scalpel/types.py
git commit -m "feat(docx-scalpel): add BlockMetadata / ListMembership / SectionInfo / NumberFormat dataclasses"
```

---

### Task 17: Add Python `DocxSession` wrapper methods

**Files:**
- Modify: `python/src/docx_scalpel/session.py`

- [ ] **Step 1: Add the new types to the import block at the top of `session.py`**

```bash
grep -n "^    AnchorInfo," python/src/docx_scalpel/session.py
```

Edit the import to add `BlockMetadata`, `ListMembership`, `SectionInfo` (alphabetized).

- [ ] **Step 2: Add the four methods after `get_anchor_infos` (around line 470)**

```python
    def get_block_metadata(self, anchor_id: str) -> BlockMetadata | None:
        """Resolve block-level metadata (style id+name, outline level, list
        membership, formatting probe) for an anchor. Returns None for unknown anchors."""
        result = self._call("get_block_metadata", {"anchorId": anchor_id})
        return BlockMetadata._from_wire(result) if result else None

    def get_block_metadatas(
        self, anchor_ids: Iterable[str]
    ) -> dict[str, BlockMetadata | None]:
        """Bulk variant of :meth:`get_block_metadata`."""
        result = self._call("get_block_metadatas", {"anchorIds": list(anchor_ids)})
        return {
            aid: BlockMetadata._from_wire(meta) if meta else None
            for aid, meta in result.items()
        }

    def get_list_membership(self, anchor_id: str) -> ListMembership | None:
        """Resolve the numbering facts for a list-item paragraph. Returns None
        when the anchor has no w:numPr."""
        result = self._call("get_list_membership", {"anchorId": anchor_id})
        return ListMembership._from_wire(result) if result else None

    def get_section_info(self, anchor_id: str) -> SectionInfo | None:
        """Resolve page-layout info for the w:sectPr that governs an anchor.
        Returns None for anchors outside the body part."""
        result = self._call("get_section_info", {"anchorId": anchor_id})
        return SectionInfo._from_wire(result) if result else None
```

- [ ] **Step 3: Verify the import is sound**

```bash
cd python && python -c "from docx_scalpel.session import DocxSession; print('ok')"
```
Expected: `ok`.

- [ ] **Step 4: Commit**

```bash
git add python/src/docx_scalpel/session.py
git commit -m "feat(docx-scalpel): expose get_block_metadata / get_list_membership / get_section_info"
```

---

### Task 18: Add Python pytest

**Files:**
- Create: `python/tests/test_block_metadata.py`

- [ ] **Step 1: Locate the fixture pattern used by existing tests**

```bash
head -40 python/tests/test_smoke.py
```

Use the same fixture loader the existing tests use (typically `conftest.py` exposes a `simple_docx` or similar fixture). If not, load bytes directly from the repo's `TestFiles/`.

- [ ] **Step 2: Create the test file**

```python
"""Coverage for the block-metadata read surface on DocxSession."""

from pathlib import Path

import pytest

from docx_scalpel import DocxSession
from docx_scalpel.types import BlockMetadata, NumberFormat


REPO_ROOT = Path(__file__).resolve().parents[2]
FIXTURE_BULLET = REPO_ROOT / "Docxodus.Tests" / "TestFiles" / "DB012-Lists-With-Different-Numberings.docx"


@pytest.fixture
def list_session() -> DocxSession:
    if not FIXTURE_BULLET.exists():
        pytest.skip(f"fixture missing: {FIXTURE_BULLET}")
    bytes_ = FIXTURE_BULLET.read_bytes()
    session = DocxSession(bytes_)
    try:
        yield session
    finally:
        session.close()


def _first_anchor_of_kind(session: DocxSession, kind: str):
    projection = session.project()
    for anchor in projection.anchor_index.values():
        if anchor.kind == kind:
            return anchor
    return None


def test_get_block_metadata_plain_paragraph(list_session: DocxSession) -> None:
    para = _first_anchor_of_kind(list_session, "p")
    if para is None:
        pytest.skip("fixture has no plain paragraph anchors")
    meta = list_session.get_block_metadata(para.id)
    assert isinstance(meta, BlockMetadata)
    assert meta.kind == "p"
    assert meta.scope == "body"


def test_get_block_metadata_unknown_anchor_returns_none(list_session: DocxSession) -> None:
    assert list_session.get_block_metadata("p:body:does-not-exist") is None


def test_get_block_metadatas_bulk_dedups(list_session: DocxSession) -> None:
    para = _first_anchor_of_kind(list_session, "p")
    if para is None:
        pytest.skip("fixture has no plain paragraph anchors")
    result = list_session.get_block_metadatas([para.id, para.id, "p:body:missing"])
    assert len(result) == 2
    assert result[para.id] is not None
    assert result["p:body:missing"] is None


def test_get_list_membership_li_anchor(list_session: DocxSession) -> None:
    li = _first_anchor_of_kind(list_session, "li")
    if li is None:
        pytest.skip("fixture has no list-item anchors")
    membership = list_session.get_list_membership(li.id)
    assert membership is not None
    assert membership.num_id > 0
    assert membership.level >= 0
    assert isinstance(membership.format, NumberFormat)


def test_get_section_info_body_anchor(list_session: DocxSession) -> None:
    para = _first_anchor_of_kind(list_session, "p")
    if para is None:
        para = _first_anchor_of_kind(list_session, "li")
    if para is None:
        pytest.skip("fixture has no body anchors")
    info = list_session.get_section_info(para.id)
    assert info is not None
    assert info.page_width_twips > 0
    assert info.columns >= 1
```

- [ ] **Step 3: Run the tests** (requires the python-host binary to be built and discoverable — same setup as existing pytest)

```bash
cd python && pytest tests/test_block_metadata.py -v 2>&1 | tail -15
```
Expected: 5 passed (or skipped if the fixture path is missing; investigate in that case).

- [ ] **Step 4: Commit**

```bash
git add python/tests/test_block_metadata.py
git commit -m "test(docx-scalpel): pytest coverage for block-metadata reads"
```

---

## Phase 9: Documentation

### Task 19: Update `CHANGELOG.md`

**Files:**
- Modify: `CHANGELOG.md`

- [ ] **Step 1: Add an entry under `[Unreleased]`** — find the section and add:

```markdown
### Added

- **`DocxSession` block-metadata read surface.** New methods
  `GetBlockMetadata` / `GetBlockMetadatas` / `GetListMembership` /
  `GetSectionInfo` expose paragraph style id+name, outline level, list
  membership (`numId`/`abstractNumId`/`ilvl`/format/start-override/
  inherited-from-style flag), and the enclosing `w:sectPr` (page
  size/orientation/margins/columns/header/footer parts). New types
  `BlockMetadata`, `ListMembership`, `SectionInfo`, and the
  `NumberFormat` enum (`Decimal`/`UpperLetter`/`LowerLetter`/
  `UpperRoman`/`LowerRoman`/`Bullet`). Surfaced in WASM
  (`DocxSessionBridge`), npm (`DocxSession.getBlockMetadata` etc.), and
  Python (`docx_scalpel.session.DocxSession.get_block_metadata` etc.).
```

- [ ] **Step 2: Commit**

```bash
git add CHANGELOG.md
git commit -m "docs(changelog): note block-metadata read surface"
```

---

### Task 20: Update `CLAUDE.md`

**Files:**
- Modify: `CLAUDE.md`

- [ ] **Step 1: Find the `DocxSession` bullet block**

```bash
grep -n "Tier C (formatting)\|Tier D (advanced)\|Tier E (annotations)" CLAUDE.md
```

- [ ] **Step 2: Insert a new bullet after the Tier E annotation entry (around line 258)**

```markdown
- Inspection: `GetBlockMetadata(anchor)`, `GetBlockMetadatas(anchors)`,
  `GetListMembership(anchor)`, `GetSectionInfo(anchor)` — read-only
  block-level metadata (style id/name, outline level, list facts:
  numId/abstractNumId/ilvl/format/start-override/from-style,
  sectPr page setup). Returns null for unknown anchors.
```

- [ ] **Step 3: Commit**

```bash
git add CLAUDE.md
git commit -m "docs(claude): document the block-metadata read surface"
```

---

### Task 21: Update `docs/architecture/docx_mutation_api.md`

**Files:**
- Modify: `docs/architecture/docx_mutation_api.md`

- [ ] **Step 1: Find the end of the doc**

```bash
tail -5 docs/architecture/docx_mutation_api.md
```

- [ ] **Step 2: Append a new "Inspection: block metadata" section at the end of the file**

```markdown
## Inspection: block metadata

`GetBlockMetadata` / `GetBlockMetadatas` / `GetListMembership` /
`GetSectionInfo` are pure reads — no mutation, no undo snapshot, no
projection invalidation. Each returns an immutable record (or null when
the anchor doesn't exist).

### `BlockMetadata`

For every block-level anchor, exposes:

- `AnchorId`, `Kind`, `Scope` — duplicated from `AnchorInfo` so the
  record is self-contained.
- `StyleId` / `StyleName` — `pStyle/@val` for paragraph kinds,
  `tblStyle/@val` for tables. `StyleName` resolves through the styles
  part's `w:name/@val`.
- `OutlineLevel` — `pPr/outlineLvl` when present; otherwise inferred
  from a `HeadingN` style (level 0..8). 0-based per Word convention.
- `List` — populated for list-item paragraphs (`null` otherwise).
- `HasInlineFormatting` — true when any descendant `w:r` carries a
  non-empty `w:rPr`. Coarse "does this paragraph have any character
  formatting at all" probe.

### `ListMembership`

For list-item paragraphs (and also surfaced as `BlockMetadata.List`):

- `NumId` / `AbstractNumId` / `Level` / `Format` — the standard
  numbering identity quadruple.
- `StartOverride` — non-null when the paragraph's `w:num` has a
  `w:lvlOverride/w:startOverride` at this level. Useful for predicting
  what `RestartNumberedList` will produce.
- `IsAutoNumbered` — always true (a paragraph without numbering returns
  `null` from `GetListMembership`).
- `FromStyle` — true when `w:numPr` is inherited from the paragraph's
  style chain (style → basedOn → basedOn → ...) rather than set inline.
  Lets callers reason about whether modifying the paragraph in place
  versus modifying the underlying style is appropriate.
- `GeneratedLabel` — same string as `AnchorInfo.AutoNumberPrefix`,
  duplicated here so callers don't take two round-trips.

### `SectionInfo`

For anchors in the body part:

- `SectionUnid` — stable id for the governing `w:sectPr`.
- `PageWidthTwips` / `PageHeightTwips` — raw twips (1 inch = 1440 twips).
- `Landscape` — true when `pgSz/@orient = "landscape"`.
- `MarginTopTwips` / `MarginBottomTwips` / `MarginLeftTwips` /
  `MarginRightTwips` — `pgMar` attribute values; defaults to 1440
  (1 inch) when missing.
- `Columns` — `cols/@num`, defaults to 1.
- `HeaderPartUris` / `FooterPartUris` — package-part URIs of the
  header/footer parts referenced via `headerReference` / `footerReference`,
  in declaration order. Empty when no headers/footers are referenced.

Returns `null` for anchors in non-body parts (footnotes, endnotes,
headers, footers, comments) — sectPr is body-only.

### `NumberFormat` enum

Closed enum used by `ListMembership.Format` (read) and by the list
write surface (when it ships). Values: `Decimal`, `UpperLetter`,
`LowerLetter`, `UpperRoman`, `LowerRoman`, `Bullet`. Any OOXML
`numFmt` value outside this set maps to `Decimal` (safest fallback).
```

- [ ] **Step 3: Commit**

```bash
git add docs/architecture/docx_mutation_api.md
git commit -m "docs(arch): add Inspection: block metadata section to docx_mutation_api"
```

---

## Phase 10: Final verification

### Task 22: Run the full test suite and prepare the PR

- [ ] **Step 1: Run all .NET tests**

```bash
dotnet test Docxodus.Tests/Docxodus.Tests.csproj 2>&1 | tail -10
```
Expected: all green (no regressions; new BM*** tests included).

- [ ] **Step 2: Run npm Playwright**

```bash
cd npm && npm run build && npm test 2>&1 | tail -10
```
Expected: all green.

- [ ] **Step 3: Run Python pytest**

```bash
cd python && pytest 2>&1 | tail -10
```
Expected: all green.

- [ ] **Step 4: Verify the branch is ready**

```bash
git log --oneline origin/main..HEAD
```
Expected: a clean linear history (~21 commits, one per task) — review for any squash candidates.

- [ ] **Step 5: Push the branch (when ready — confirm with the user before pushing)**

```bash
git push -u origin feat/docx-block-inspection
```

- [ ] **Step 6: Open the PR (after the user confirms)**

```bash
gh pr create --title "feat(session): block metadata read surface (GetBlockMetadata/GetListMembership/GetSectionInfo)" --body "$(cat <<'EOF'
## Summary
- Adds read-only block-metadata methods to `DocxSession`: `GetBlockMetadata`, `GetBlockMetadatas`, `GetListMembership`, `GetSectionInfo`.
- New types: `BlockMetadata`, `ListMembership`, `SectionInfo`, `NumberFormat` enum.
- Routed through `DocxSessionOps` → WASM `[JSExport]` → stdio dispatcher → npm `DocxSession` → python `DocxSession`.

## Test plan
- [ ] `dotnet test` (10 new BM tests, no regressions)
- [ ] `npm run build && npm test` (5 new Playwright specs)
- [ ] `pytest python/tests/test_block_metadata.py` (5 new pytest cases)
- [ ] Round-trip a 100-page doc through `GetBlockMetadata`/`GetListMembership`/`GetSectionInfo` to confirm <5ms per call typical
EOF
)"
```

---

## Spec coverage cross-check

This plan implements exactly the **block-metadata** sub-feature of the spec
(`docs/superpowers/specs/2026-05-28-note-refs-list-writes-metadata-design.md`,
Feature area 3). The other two sub-features get their own plans:

- **List writes** (Feature area 2) — to be written when this PR lands.
- **Note refs** (Feature area 1) — to be written last per the spec's
  implementation order.

Each of those is independently testable software on its own. The
`NumberFormat` enum and `ListMembership` type added here are reused by
the list-writes plan; tests there will round-trip writes through these
reads to verify behavior.
