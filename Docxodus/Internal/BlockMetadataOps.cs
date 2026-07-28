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
        // WmlToHtmlConverter for the same usage pattern. Twips are integer-valued.
        int width = ParseInt((string?)pgSz?.Attribute(W._w)) ?? 12240;   // 8.5"
        int height = ParseInt((string?)pgSz?.Attribute(W.h)) ?? 15840;  // 11"
        bool landscape = string.Equals((string?)pgSz?.Attribute(W.orient), "landscape",
            System.StringComparison.Ordinal);

        int top = ParseInt((string?)pgMar?.Attribute(W.top)) ?? 1440;
        int bottom = ParseInt((string?)pgMar?.Attribute(W.bottom)) ?? 1440;
        int left = ParseInt((string?)pgMar?.Attribute(W.left)) ?? 1440;
        int right = ParseInt((string?)pgMar?.Attribute(W.right)) ?? 1440;

        int colCount = 1;
        if (cols is not null && int.TryParse((string?)cols.Attribute(W.num), out var parsedCols))
            colCount = parsedCols;

        var (ownHeaderRefs, ownFooterRefs) = ResolveSectionHeaderFooterRefs(doc, sectPr);
        var (headerRefs, footerRefs) = AddInheritedRefs(doc, element, sectPr, ownHeaderRefs, ownFooterRefs);

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
            // The URI lists keep their original meaning — this section's OWN references — so only
            // the (new) *Refs lists carry inherited stories.
            HeaderPartUris = ownHeaderRefs.Select(r => r.PartUri).ToList(),
            FooterPartUris = ownFooterRefs.Select(r => r.PartUri).ToList(),
            HeaderRefs = headerRefs,
            FooterRefs = footerRefs,
        };
    }

    /// <summary>
    /// The <c>w:sectPr</c> that governs <paramref name="element"/> (a body block): the next
    /// forward paragraph's <c>pPr/sectPr</c> (a mid-document section break), else the element's
    /// own <c>pPr/sectPr</c>, else the body's trailing <c>sectPr</c> (the document-final section).
    /// Returns <c>null</c> when the body has no section properties at all. Shared with
    /// <see cref="DocxSession.SetHeaderText"/>/<see cref="DocxSession.SetFooterText"/>.
    /// </summary>
    internal static XElement? FindGoverningSectPr(XElement element)
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

    /// <summary>
    /// Extend a section's own references with the ones it INHERITS. Per ECMA-376 §17.6.17 a
    /// section that declares no reference of a given type continues the nearest preceding
    /// section's, which is why a multi-section document commonly defines its headers once, in
    /// the first section, and leaves the rest empty. Reporting only a section's own references
    /// would tell a caller "this part of the document has no header" when it visibly does — and
    /// an editor acting on that would mint a redundant part and break the inheritance.
    /// </summary>
    private static (IReadOnlyList<HeaderFooterRef> headers, IReadOnlyList<HeaderFooterRef> footers)
        AddInheritedRefs(WordprocessingDocument doc, XElement element, XElement governing,
            IReadOnlyList<HeaderFooterRef> ownHeaders, IReadOnlyList<HeaderFooterRef> ownFooters)
    {
        var body = element.AncestorsAndSelf(W.body).FirstOrDefault();
        if (body is null) return (ownHeaders, ownFooters);

        // Every sectPr in document order — inline section breaks (w:pPr/w:sectPr) followed by the
        // body's trailing one, which is exactly the order sections apply in.
        var all = body.Descendants(W.sectPr).ToList();
        int idx = all.IndexOf(governing);
        if (idx <= 0) return (ownHeaders, ownFooters);

        var headers = ownHeaders.ToList();
        var footers = ownFooters.ToList();
        for (int i = idx - 1; i >= 0 && (headers.Count < 3 || footers.Count < 3); i--)
        {
            var (prevHeaders, prevFooters) = ResolveSectionHeaderFooterRefs(doc, all[i]);
            foreach (var r in prevHeaders)
                if (!headers.Any(h => h.Kind == r.Kind))
                    headers.Add(r with { Inherited = true });
            foreach (var r in prevFooters)
                if (!footers.Any(f => f.Kind == r.Kind))
                    footers.Add(r with { Inherited = true });
        }
        return (headers, footers);
    }

    /// <summary>The story kind a <c>w:headerReference</c>/<c>w:footerReference</c> supplies.
    /// <c>w:type</c> is optional; an absent or unrecognized value means the default story.</summary>
    private static HeaderFooterKind ParseRefKind(XElement reference)
        => (string?)reference.Attribute(W.type) switch
        {
            "first" => HeaderFooterKind.First,
            "even" => HeaderFooterKind.Even,
            _ => HeaderFooterKind.Default,
        };

    private static (IReadOnlyList<HeaderFooterRef> headers, IReadOnlyList<HeaderFooterRef> footers)
        ResolveSectionHeaderFooterRefs(WordprocessingDocument doc, XElement sectPr)
    {
        var main = doc.MainDocumentPart;
        if (main is null)
            return (System.Array.Empty<HeaderFooterRef>(), System.Array.Empty<HeaderFooterRef>());

        var headers = new List<HeaderFooterRef>();
        var footers = new List<HeaderFooterRef>();

        foreach (var headerRef in sectPr.Elements(W.headerReference))
        {
            var rId = (string?)headerRef.Attribute(R.id);
            if (rId is null) continue;
            if (main.GetPartById(rId) is HeaderPart part)
                headers.Add(new HeaderFooterRef
                {
                    Kind = ParseRefKind(headerRef),
                    PartUri = part.Uri.ToString(),
                });
        }
        foreach (var footerRef in sectPr.Elements(W.footerReference))
        {
            var rId = (string?)footerRef.Attribute(R.id);
            if (rId is null) continue;
            if (main.GetPartById(rId) is FooterPart part)
                footers.Add(new HeaderFooterRef
                {
                    Kind = ParseRefKind(footerRef),
                    PartUri = part.Uri.ToString(),
                });
        }
        return (headers, footers);
    }

    private static int? ParseInt(string? raw)
        => int.TryParse(raw, System.Globalization.NumberStyles.Any,
            System.Globalization.CultureInfo.InvariantCulture, out var v) ? v : null;

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

    private static bool HasInlineFormatting(XElement element)
    {
        return element.Descendants(W.r).Any(r =>
        {
            var rPr = r.Element(W.rPr);
            return rPr is not null && rPr.Elements().Any();
        });
    }
}
