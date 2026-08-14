#nullable enable

// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus.Internal;

/// <summary>
/// Pure style and formatting resolvers for the session inspection surface. The effective paths
/// deliberately route through <see cref="FormattingAssembler"/>'s style rollups so inspection and
/// rendering do not grow independent inheritance implementations.
/// </summary>
internal static class FormattingIntrospectionOps
{
    public static IReadOnlyList<StyleInfo> ListStyles(WordprocessingDocument doc)
    {
        var root = doc.MainDocumentPart?.StyleDefinitionsPart?.GetXDocument().Root;
        if (root is null) return Array.Empty<StyleInfo>();

        var latent = root.Element(W.latentStyles);
        var result = new List<StyleInfo>();
        foreach (var style in root.Elements(W.style))
        {
            var id = (string?)style.Attribute(W.styleId);
            if (string.IsNullOrEmpty(id)) continue;
            var name = (string?)style.Element(W.name)?.Attribute(W.val) ?? id;
            var type = (string?)style.Attribute(W.type) ?? "paragraph";
            var latentException = latent?.Elements(W.lsdException).FirstOrDefault(e =>
                string.Equals((string?)e.Attribute(W.name), name, StringComparison.OrdinalIgnoreCase));

            ParagraphFormatting? paragraph = null;
            RunFormattingInfo? run = null;
            TableStyleFormatting? table = null;

            if (type == "paragraph")
            {
                var synthetic = SyntheticParagraph(id, characterStyleId: null);
                paragraph = ParseParagraph(
                    FormattingAssembler.ResolveEffectiveParagraphProperties(doc, synthetic),
                    effective: true,
                    effectiveStyleId: id);
                run = ParseRun(
                    FormattingAssembler.ResolveEffectiveRunProperties(doc, synthetic.Element(W.r)!),
                    effective: true,
                    effectiveStyleId: null);
            }
            else if (type == "character")
            {
                var synthetic = SyntheticParagraph(paragraphStyleId: null, characterStyleId: id);
                run = ParseRun(
                    FormattingAssembler.ResolveEffectiveRunProperties(doc, synthetic.Element(W.r)!),
                    effective: true,
                    effectiveStyleId: id);
            }
            else if (type == "table")
            {
                table = ParseTableStyle(FormattingAssembler.ResolveTableStyle(doc, id));
            }

            result.Add(new StyleInfo
            {
                Id = id,
                Name = name,
                Type = type,
                BasedOn = (string?)style.Element(W.basedOn)?.Attribute(W.val),
                Next = (string?)style.Element(W.next)?.Attribute(W.val),
                IsDefault = ReadOnOffAttribute(style.Attribute(W._default)) == true,
                IsCustom = ReadOnOffAttribute(style.Attribute(W.customStyle)) == true,
                HasLatentException = latentException is not null,
                UiPriority = ReadIntAttribute(style.Element(W.uiPriority)?.Attribute(W.val))
                    ?? ReadIntAttribute(latentException?.Attribute(W.uiPriority))
                    ?? ReadIntAttribute(latent?.Attribute(W.defUIPriority)),
                SemiHidden = ResolveGalleryBool(style, W.semiHidden,
                    latentException, W.semiHidden, latent, W.defSemiHidden),
                UnhideWhenUsed = ResolveGalleryBool(style, W.unhideWhenUsed,
                    latentException, W.unhideWhenUsed, latent, W.defUnhideWhenUsed),
                QuickFormat = ResolveGalleryBool(style, W.qFormat,
                    latentException, W.qFormat, latent, W.defQFormat),
                Locked = ResolveGalleryBool(style, W.locked,
                    latentException, W.locked, latent, W.defLockedState),
                ResolvedParagraph = paragraph,
                ResolvedRun = run,
                ResolvedTable = table,
            });
        }
        return result;
    }

    public static FormattingInspection? GetFormatting(WordprocessingDocument doc, AnchorTarget target)
    {
        var paragraph = target.Resolve(doc);
        if (paragraph is null || paragraph.Name != W.p) return null;

        var direct = ParseParagraph(paragraph.Element(W.pPr), effective: false, effectiveStyleId: null);
        var effectiveStyleId = (string?)paragraph.Element(W.pPr)?.Element(W.pStyle)?.Attribute(W.val)
            ?? DefaultParagraphStyleId(doc);
        var effective = ParseParagraph(
            FormattingAssembler.ResolveEffectiveParagraphProperties(doc, paragraph),
            effective: true,
            effectiveStyleId);

        return new FormattingInspection
        {
            AnchorId = target.Anchor.Id,
            DirectParagraph = direct,
            EffectiveParagraph = effective,
            Runs = ListInlineSpans(doc, target),
        };
    }

    public static IReadOnlyList<InlineSpan> ListInlineSpans(
        WordprocessingDocument doc, AnchorTarget target)
    {
        var paragraph = target.Resolve(doc);
        if (paragraph is null || paragraph.Name != W.p) return Array.Empty<InlineSpan>();

        var map = RunTextMap.Build(paragraph);
        var result = new List<InlineSpan>(map.Segments.Count);
        foreach (var segment in map.Segments)
        {
            var run = segment.Run;
            // The anchor index normally assigned this already. For unusual/custom projection
            // scopes, derive the same deterministic identity without mutating the live package.
            var unid = UnidHelper.ReadOrDeriveUnid(run);

            var directStyleId = (string?)run.Element(W.rPr)?.Element(W.rStyle)?.Attribute(W.val);
            result.Add(new InlineSpan
            {
                AnchorId = target.Anchor.Id,
                RunUnid = unid,
                Span = new CharSpan(segment.StartOffsetInBlock, segment.Length),
                Text = DocxSession.RunText(run),
                Direct = ParseRun(run.Element(W.rPr), effective: false, effectiveStyleId: null),
                Effective = ParseRun(
                    FormattingAssembler.ResolveEffectiveRunProperties(doc, run),
                    effective: true,
                    effectiveStyleId: directStyleId),
                ContentControlAnchorIds = run.Ancestors(W.sdt).Reverse()
                    .Select(control => (string?)control.Attribute(PtOpenXml.Unid))
                    .Where(unid => !string.IsNullOrEmpty(unid))
                    .Select(unid => $"sdt:{target.Anchor.Scope}:{unid}")
                    .ToArray(),
            });
        }
        return result;
    }

    private static XElement SyntheticParagraph(string? paragraphStyleId, string? characterStyleId)
    {
        var pPr = new XElement(W.pPr,
            paragraphStyleId is null
                ? null
                : new XElement(W.pStyle, new XAttribute(W.val, paragraphStyleId)));
        var rPr = new XElement(W.rPr,
            characterStyleId is null
                ? null
                : new XElement(W.rStyle, new XAttribute(W.val, characterStyleId)));
        return new XElement(W.p, pPr,
            new XElement(W.r, rPr, new XElement(W.t, "x")));
    }

    private static string? DefaultParagraphStyleId(WordprocessingDocument doc) =>
        (string?)doc.MainDocumentPart?.StyleDefinitionsPart?.GetXDocument().Root?
            .Elements(W.style)
            .FirstOrDefault(s => (string?)s.Attribute(W.type) == "paragraph"
                && ReadOnOffAttribute(s.Attribute(W._default)) == true)?
            .Attribute(W.styleId);

    private static ParagraphFormatting ParseParagraph(
        XElement? pPr, bool effective, string? effectiveStyleId)
    {
        var ind = pPr?.Element(W.ind);
        var spacing = pPr?.Element(W.spacing);
        var alignment = ParseAlignment((string?)pPr?.Element(W.jc)?.Attribute(W.val));
        var line = ReadIntAttribute(spacing?.Attribute(W.line));
        var lineRule = ParseLineSpacingRule((string?)spacing?.Attribute(W.lineRule));

        var value = new ParagraphFormatting
        {
            StyleId = effectiveStyleId
                ?? (string?)pPr?.Element(W.pStyle)?.Attribute(W.val),
            Alignment = alignment,
            LeftIndentTwips = ReadIntAttribute(ind?.Attribute(W.left) ?? ind?.Attribute(W.start)),
            RightIndentTwips = ReadIntAttribute(ind?.Attribute(W.right) ?? ind?.Attribute(W.end)),
            FirstLineIndentTwips = ReadIntAttribute(ind?.Attribute(W.firstLine)),
            HangingIndentTwips = ReadIntAttribute(ind?.Attribute(W.hanging)),
            SpacingBeforeTwips = ReadIntAttribute(spacing?.Attribute(W.before)),
            SpacingAfterTwips = ReadIntAttribute(spacing?.Attribute(W.after)),
            LineSpacing = line,
            LineSpacingRule = line is null ? null : lineRule ?? LineSpacingRule.Auto,
            KeepNext = ReadOnOffElement(pPr?.Element(W.keepNext)),
            KeepLines = ReadOnOffElement(pPr?.Element(W.keepLines)),
            PageBreakBefore = ReadOnOffElement(pPr?.Element(W.pageBreakBefore)),
            OutlineLevel = ReadIntAttribute(pPr?.Element(W.outlineLvl)?.Attribute(W.val)),
            ShadingFill = (string?)pPr?.Element(W.shd)?.Attribute(W.fill),
            TopBorder = ParseBorder(pPr?.Element(W.pBdr)?.Element(W.top)),
            BottomBorder = ParseBorder(pPr?.Element(W.pBdr)?.Element(W.bottom)),
        };

        if (!effective) return value;
        return value with
        {
            Alignment = value.Alignment ?? ParagraphAlignment.Left,
            LeftIndentTwips = value.LeftIndentTwips ?? 0,
            RightIndentTwips = value.RightIndentTwips ?? 0,
            FirstLineIndentTwips = value.FirstLineIndentTwips ?? 0,
            HangingIndentTwips = value.HangingIndentTwips ?? 0,
            SpacingBeforeTwips = value.SpacingBeforeTwips ?? 0,
            SpacingAfterTwips = value.SpacingAfterTwips ?? 0,
            LineSpacing = value.LineSpacing ?? 240,
            LineSpacingRule = value.LineSpacingRule ?? LineSpacingRule.Auto,
            KeepNext = value.KeepNext ?? false,
            KeepLines = value.KeepLines ?? false,
            PageBreakBefore = value.PageBreakBefore ?? false,
        };
    }

    private static RunFormattingInfo ParseRun(
        XElement? rPr, bool effective, string? effectiveStyleId)
    {
        var styleId = effectiveStyleId
            ?? (string?)rPr?.Element(W.rStyle)?.Attribute(W.val);
        var underlineElement = rPr?.Element(W.u);
        var underlineStyle = (string?)underlineElement?.Attribute(W.val);
        bool? underline = underlineElement is null
            ? null
            : !string.Equals(underlineStyle, "none", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(underlineStyle, "0", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(underlineStyle, "false", StringComparison.OrdinalIgnoreCase);
        if (underline == true && string.IsNullOrEmpty(underlineStyle)) underlineStyle = "single";

        var sizeHalfPoints = ReadIntAttribute(rPr?.Element(W.sz)?.Attribute(W.val));
        var fonts = rPr?.Element(W.rFonts);
        var value = new RunFormattingInfo
        {
            StyleId = styleId,
            Bold = ReadOnOffElement(rPr?.Element(W.b)),
            Italic = ReadOnOffElement(rPr?.Element(W.i)),
            Underline = underline,
            UnderlineStyle = underlineStyle,
            Strike = ReadOnOffElement(rPr?.Element(W.strike))
                ?? ReadOnOffElement(rPr?.Element(W.dstrike)),
            Code = styleId is null ? null : string.Equals(styleId, "Code", StringComparison.Ordinal),
            Color = (string?)rPr?.Element(W.color)?.Attribute(W.val),
            Highlight = (string?)rPr?.Element(W.highlight)?.Attribute(W.val),
            VertAlign = (string?)rPr?.Element(W.vertAlign)?.Attribute(W.val),
            FontSizePts = sizeHalfPoints is null ? null : sizeHalfPoints.Value / 2.0,
            FontFamily = (string?)fonts?.Attribute(W.ascii)
                ?? (string?)fonts?.Attribute(W.hAnsi)
                ?? (string?)fonts?.Attribute(W.cs),
            Caps = ReadOnOffElement(rPr?.Element(W.caps)),
            SmallCaps = ReadOnOffElement(rPr?.Element(W.smallCaps)),
            Hidden = ReadOnOffElement(rPr?.Element(W.vanish)),
        };

        if (!effective) return value;
        return value with
        {
            Bold = value.Bold ?? false,
            Italic = value.Italic ?? false,
            Underline = value.Underline ?? false,
            Strike = value.Strike ?? false,
            Code = value.Code ?? false,
            Caps = value.Caps ?? false,
            SmallCaps = value.SmallCaps ?? false,
            Hidden = value.Hidden ?? false,
        };
    }

    private static TableStyleFormatting ParseTableStyle(XElement style)
    {
        var tblPr = style.Element(W.tblPr);
        var width = tblPr?.Element(W.tblW);
        var indent = tblPr?.Element(W.tblInd);
        var widthTwips = string.Equals((string?)width?.Attribute(W.type), "dxa", StringComparison.Ordinal)
            ? ReadIntAttribute(width?.Attribute(W._w))
            : null;
        var indentTwips = string.Equals((string?)indent?.Attribute(W.type), "dxa", StringComparison.Ordinal)
            ? ReadIntAttribute(indent?.Attribute(W._w))
            : null;
        return new TableStyleFormatting
        {
            Alignment = (string?)tblPr?.Element(W.jc)?.Attribute(W.val),
            WidthTwips = widthTwips,
            IndentTwips = indentTwips,
            Layout = (string?)tblPr?.Element(W.tblLayout)?.Attribute(W.type),
            HasBorders = tblPr?.Element(W.tblBorders)?.Elements().Any(),
            CellShadingFill = (string?)style.Element(W.tcPr)?.Element(W.shd)?.Attribute(W.fill),
        };
    }

    private static ParagraphBorderEdge? ParseBorder(XElement? edge)
    {
        if (edge is null) return null;
        return new ParagraphBorderEdge
        {
            Style = (string?)edge.Attribute(W.val),
            Size = ReadIntAttribute(edge.Attribute(W.sz)),
            Color = (string?)edge.Attribute(W.color),
            Space = ReadIntAttribute(edge.Attribute(W.space)),
        };
    }

    private static ParagraphAlignment? ParseAlignment(string? raw) => raw switch
    {
        "left" or "start" => ParagraphAlignment.Left,
        "center" => ParagraphAlignment.Center,
        "right" or "end" => ParagraphAlignment.Right,
        "both" or "distribute" => ParagraphAlignment.Justify,
        _ => null,
    };

    private static LineSpacingRule? ParseLineSpacingRule(string? raw) => raw switch
    {
        "auto" => LineSpacingRule.Auto,
        "exact" => LineSpacingRule.Exact,
        "atLeast" => LineSpacingRule.AtLeast,
        _ => null,
    };

    private static bool? ResolveGalleryBool(
        XElement style, XName styleName,
        XElement? exception, XName exceptionName, XElement? latent, XName defaultName) =>
        ReadOnOffElement(style.Element(styleName))
            ?? ReadOnOffAttribute(exception?.Attribute(exceptionName))
            ?? ReadOnOffAttribute(latent?.Attribute(defaultName));

    private static int? ReadIntAttribute(XAttribute? attribute) =>
        int.TryParse((string?)attribute, NumberStyles.Integer, CultureInfo.InvariantCulture, out var value)
            ? value
            : null;

    private static bool? ReadOnOffElement(XElement? element) =>
        element is null ? null : ReadOnOffAttribute(element.Attribute(W.val)) ?? true;

    private static bool? ReadOnOffAttribute(XAttribute? attribute)
    {
        if (attribute is null) return null;
        return attribute.Value switch
        {
            "1" or "true" or "on" => true,
            "0" or "false" or "off" => false,
            _ => null,
        };
    }
}
