// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus;

/// <summary>How text lines up at a tab stop (<c>w:tab/@w:val</c>). A bar tab draws a vertical
/// line at its position instead of stopping text there.</summary>
public enum TabStopAlignment { Left, Center, Right, Decimal, Bar }

/// <summary>The characters that fill the space before a tab stop (<c>w:tab/@w:leader</c>).</summary>
public enum TabStopLeader { None, Dot, Hyphen, Underscore, Heavy, MiddleDot }

/// <summary>One tab stop: its position from the paragraph's leading text edge (the left margin in
/// a left-to-right paragraph, not its indent), in twips, its alignment and its leader.</summary>
public sealed record ParagraphTabStop
{
    public required int PositionTwips { get; init; }
    public TabStopAlignment Alignment { get; init; } = TabStopAlignment.Left;
    public TabStopLeader Leader { get; init; } = TabStopLeader.None;
}

/// <summary>What <see cref="DocxSession.GetTabStops"/> reads: the tab stops in effect for a
/// paragraph (its style chain and its own, with cleared ones removed), the ones it would have
/// without its own <c>w:tabs</c>, and the document's default tab interval.</summary>
public sealed record TabStopsInfo
{
    /// <summary>The stops in effect, by position.</summary>
    public required IReadOnlyList<ParagraphTabStop> Stops { get; init; }
    /// <summary>The stops the paragraph's styles give it, by position.</summary>
    public required IReadOnlyList<ParagraphTabStop> StyleStops { get; init; }
    /// <summary><c>w:defaultTabStop</c> from settings, twips (720, half an inch, when absent).</summary>
    public int DefaultTabStopTwips { get; init; }
}

public sealed partial class DocxSession
{
    /// <summary>Word's default tab interval when settings has no <c>w:defaultTabStop</c>.</summary>
    private const int WordDefaultTabStopTwips = 720;

    /// <summary>The largest tab position Word accepts (22 inches), either side of the edge.</summary>
    private const int MaxTabPositionTwips = 31680;

    /// <summary>
    /// The tab stops of the paragraph <paramref name="anchorId"/> names, as Word's Tabs dialog and
    /// ruler show them; null when the anchor is not a paragraph. Numbering-level tabs are not
    /// included (as for <see cref="GetFormatting"/>'s effective paragraph).
    /// </summary>
    public TabStopsInfo? GetTabStops(string anchorId)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(anchorId);
        var target = FindAnchor(anchorId);
        if (target is null || target.Anchor.Kind is not ("p" or "h" or "li")) return null;
        var paragraph = target.Resolve(_doc!);
        if (paragraph?.Name != W.p) return null;
        return new TabStopsInfo
        {
            Stops = ReadTabStops(FormattingAssembler.ResolveEffectiveParagraphProperties(_doc!, paragraph)
                .Element(W.tabs)),
            StyleStops = InheritedTabStops(paragraph),
            DefaultTabStopTwips = DefaultTabStopTwips(),
        };
    }

    /// <summary>
    /// Make <paramref name="stops"/> the tab stops in effect for the paragraph
    /// <paramref name="anchorId"/> names, as Word does: its own <c>w:tabs</c> holds each stop its
    /// styles do not already give it, and a <c>clear</c> entry for each style stop that is not
    /// wanted. An empty list clears every stop (default tabs remain). One undo step; with changes
    /// tracked, one <c>w:pPrChange</c>.
    /// </summary>
    public EditResult SetTabStops(string anchorId, IReadOnlyList<ParagraphTabStop> stops)
    {
        ArgumentNullException.ThrowIfNull(stops);
        if (MutationRefusal() is { } refusal) return EditResult.Fail(refusal);
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "anchor not found", anchorId);
        if (target.Anchor.Kind is not ("p" or "h" or "li"))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind, "SetTabStops requires a paragraph anchor", anchorId);
        var element = target.Resolve(_doc!);
        if (element is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "element null", anchorId);
        if (stops.Any(s => s is null || Math.Abs(s.PositionTwips) > MaxTabPositionTwips))
            return EditResult.Fail(EditErrorCode.InvalidParagraphFormat,
                $"tab positions must be within ±{MaxTabPositionTwips} twips", anchorId);
        if (stops.GroupBy(s => s.PositionTwips).Any(g => g.Count() > 1))
            return EditResult.Fail(EditErrorCode.InvalidParagraphFormat,
                "two tab stops cannot share a position", anchorId);
        if (RefuseNestedTrackedParagraphPropertyChange(element, anchorId) is { } pending) return pending;

        var inherited = InheritedTabStops(element);
        var own = new List<XElement>();
        foreach (var stop in stops)
            if (!inherited.Contains(stop))
                own.Add(TabElement(AlignmentValue(stop.Alignment), stop.PositionTwips, stop.Leader));
        foreach (var stop in inherited)
            if (!stops.Any(s => s.PositionTwips == stop.PositionTwips))
                own.Add(TabElement("clear", stop.PositionTwips, TabStopLeader.None));
        own.Sort((a, b) => ((int)a.Attribute(W.pos)!).CompareTo((int)b.Attribute(W.pos)!));

        var oldPPr = new XElement(element.Element(W.pPr) ?? new XElement(W.pPr));
        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var pPr = element.Element(W.pPr);
            if (pPr is null) { pPr = new XElement(W.pPr); element.AddFirst(pPr); }
            pPr.Elements(W.tabs).Remove();
            if (own.Count > 0) SetPPrChildInOrder(pPr, new XElement(W.tabs, own));

            if (_trackedChanges == TrackedChangeMode.RenderInline
                && !PropertySnapshotEquals(oldPPr, pPr, W.pPrChange, W.rPr, W.sectPr))
                TrackPropertyMutation(pPr, oldPPr, W.pPrChange,
                    _revisionAuthor ?? "docxodus", NextTrackedFormatRevisionDate(), W.rPr, W.sectPr);

            // A pPr-only write: the anchor's kind, scope and unid are unchanged (see SetParagraphFormat).
            InvalidateProjectionCache();
            return new EditResult { Success = true, Modified = new[] { target.Anchor }, Patch = PatchFor(target) };
        }
        catch (Exception ex)
        {
            return FailInternal(ex, anchorId);
        }
    }

    /// <summary>
    /// Set the document's default tab interval (<c>w:defaultTabStop</c> in settings), the spacing
    /// of the automatic stops after the last set one. One undo step. Not tracked: Word records no
    /// revision for a settings change.
    /// </summary>
    public EditResult SetDefaultTabStop(int twips)
    {
        if (MutationRefusal() is { } refusal) return EditResult.Fail(refusal);
        if (twips is <= 0 or > MaxTabPositionTwips)
            return EditResult.Fail(EditErrorCode.InvalidParagraphFormat,
                $"the default tab stop must be between 1 and {MaxTabPositionTwips} twips", null);
        var main = _doc!.MainDocumentPart;
        if (main is null) return EditResult.Fail(EditErrorCode.InternalError, "no main document part", null);
        if (DefaultTabStopTwips() == twips && main.DocumentSettingsPart?.GetXDocument().Root?.Element(W.defaultTabStop) is not null)
            return new EditResult { Success = true };

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var settingsPart = main.DocumentSettingsPart ?? main.AddNewPart<DocumentSettingsPart>();
            var xDoc = settingsPart.GetXDocument();
            var root = xDoc.Root;
            if (root is null)
            {
                root = new XElement(W.settings, new XAttribute(XNamespace.Xmlns + "w", W.w));
                xDoc.Add(root);
            }
            root.Elements(W.defaultTabStop).Remove();
            WordprocessingMLUtil.EnsureSettingsChildInOrder(root,
                new XElement(W.defaultTabStop, new XAttribute(W.val, twips)));
            settingsPart.PutXDocument();
            InvalidateProjectionCache();
            return new EditResult { Success = true };
        }
        catch (Exception ex)
        {
            return FailInternal(ex, null);
        }
    }

    /// <summary>The stops the paragraph's styles give it: its effective tabs without its own w:tabs.</summary>
    private IReadOnlyList<ParagraphTabStop> InheritedTabStops(XElement paragraph)
    {
        var pPr = new XElement(paragraph.Element(W.pPr) ?? new XElement(W.pPr));
        pPr.Elements(W.tabs).Remove();
        var probe = new XElement(W.p, paragraph.Attributes(), pPr);
        return ReadTabStops(FormattingAssembler.ResolveEffectiveParagraphProperties(_doc!, probe)
            .Element(W.tabs));
    }

    private int DefaultTabStopTwips()
    {
        var value = (string?)_doc!.MainDocumentPart?.DocumentSettingsPart?.GetXDocument().Root?
            .Element(W.defaultTabStop)?.Attribute(W.val);
        return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) && n > 0
            ? n : WordDefaultTabStopTwips;
    }

    /// <summary>The set stops of a merged w:tabs (clear entries dropped), by position. A
    /// numbering tab (<c>num</c>) reads as a left stop.</summary>
    private static IReadOnlyList<ParagraphTabStop> ReadTabStops(XElement? tabs)
    {
        if (tabs is null) return Array.Empty<ParagraphTabStop>();
        var stops = new List<ParagraphTabStop>();
        foreach (var tab in tabs.Elements(W.tab))
        {
            var val = (string?)tab.Attribute(W.val);
            if (val == "clear") continue;
            var pos = (string?)tab.Attribute(W.pos);
            if (pos is null) continue;
            stops.Add(new ParagraphTabStop
            {
                PositionTwips = WordprocessingMLUtil.StringToTwips(pos),
                Alignment = val switch
                {
                    "center" => TabStopAlignment.Center,
                    "right" or "end" => TabStopAlignment.Right,
                    "decimal" => TabStopAlignment.Decimal,
                    "bar" => TabStopAlignment.Bar,
                    _ => TabStopAlignment.Left,
                },
                Leader = (string?)tab.Attribute(W.leader) switch
                {
                    "dot" => TabStopLeader.Dot,
                    "hyphen" => TabStopLeader.Hyphen,
                    "underscore" => TabStopLeader.Underscore,
                    "heavy" => TabStopLeader.Heavy,
                    "middleDot" => TabStopLeader.MiddleDot,
                    _ => TabStopLeader.None,
                },
            });
        }
        return stops.GroupBy(s => s.PositionTwips).Select(g => g.First()).OrderBy(s => s.PositionTwips).ToList();
    }

    private static string AlignmentValue(TabStopAlignment alignment) => alignment switch
    {
        TabStopAlignment.Center => "center",
        TabStopAlignment.Right => "right",
        TabStopAlignment.Decimal => "decimal",
        TabStopAlignment.Bar => "bar",
        _ => "left",
    };

    private static XElement TabElement(string val, int pos, TabStopLeader leader) =>
        new(W.tab,
            new XAttribute(W.val, val),
            leader == TabStopLeader.None ? null : new XAttribute(W.leader, leader switch
            {
                TabStopLeader.Dot => "dot",
                TabStopLeader.Hyphen => "hyphen",
                TabStopLeader.Underscore => "underscore",
                TabStopLeader.Heavy => "heavy",
                _ => "middleDot",
            }),
            new XAttribute(W.pos, pos));
}
