using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace Docxodus.Ir.Diff;

internal static partial class IrMarkupRenderer
{
    private static readonly XName[] MarginAxes = { W.top, W.left, W.bottom, W.right, W.start, W.end };

    /// <summary>
    /// Table-style geometry has no style-level native property history. Keep the left style graph
    /// and express its margin deltas on the affected right table/cell sources instead. Resolve
    /// conditional regions through the normal assembler, which also preserves partial overrides.
    /// Existing paired shell emitters replace these provisional archives with the true left shell.
    /// </summary>
    private static IrDocument ProjectTableStyleMargins(WmlDocument left, WmlDocument right, RenderState state)
    {
        if (state.Settings.PreserveInputRevisions || !state.Settings.TrackTableFormatChanges ||
            !state.RightSource.Sources.Values.Any(x => x.Descendants(W.tbl).Any()))
            return state.RightSource;
        using var leftStream = new OpenXmlMemoryStreamDocument(left);
        using var rightStream = new OpenXmlMemoryStreamDocument(right);
        using var leftDoc = leftStream.GetWordprocessingDocument();
        using var rightDoc = rightStream.GetWordprocessingDocument();
        var ls = leftDoc.MainDocumentPart?.StyleDefinitionsPart?.GetXDocument().Root;
        var rs = rightDoc.MainDocumentPart?.StyleDefinitionsPart?.GetXDocument().Root;
        if (ls is null || rs is null || XNode.DeepEquals(MarginDefinitions(ls), MarginDefinitions(rs)))
            return state.RightSource;

        // A chain needs projection when it reaches a retained left definition. A right-only
        // child is imported, but still inherits that left ancestor's old margins in the output.
        var used = state.RightSource.Sources.Values.SelectMany(x => x.Descendants(W.tblStyle))
            .Select(e => (string?)e.Attribute(W.val)).Where(id => id is not null).Cast<string>().ToHashSet(StringComparer.Ordinal);
        used.RemoveWhere(id => !UsesRetainedTableStyle(ls, rs, id));
        if (used.Count == 0)
            return state.RightSource;

        var inherited = AssembleConsumers(right, state.RightSource, null, replaceDefaults: false, styles =>
        {
            var leftIds = ls.Elements(W.style).Where(s => (string?)s.Attribute(W.type) == "table")
                .Select(s => (string?)s.Attribute(W.styleId)).ToHashSet(StringComparer.Ordinal);
            styles.Elements(W.style).Where(s => (string?)s.Attribute(W.type) == "table" && leftIds.Contains((string?)s.Attribute(W.styleId))).Remove();
            styles.Add(ls.Elements(W.style).Where(s => (string?)s.Attribute(W.type) == "table").Select(s => new XElement(s)));
        });
        var revised = AssembleConsumers(right, state.RightSource, null, replaceDefaults: false);
        var copies = new Dictionary<XElement, XElement>();
        var sources = new Dictionary<Uri, XDocument>();
        foreach (var (uri, source) in state.RightSource.Sources)
        {
            var clone = new XDocument(source);
            sources.Add(uri, clone);
            foreach (var pair in source.Descendants().Zip(clone.Descendants()))
                copies.Add(pair.First, pair.Second);
            if (!inherited.TryGetValue(uri, out var oldEffective) || !revised.TryGetValue(uri, out var newEffective))
                continue;
            var tables = source.Descendants(W.tbl).ToArray();
            var oldTables = oldEffective.Descendants(W.tbl).ToArray();
            var newTables = newEffective.Descendants(W.tbl).ToArray();
            if (tables.Length != oldTables.Length || tables.Length != newTables.Length)
                return state.RightSource;
            for (int i = 0; i < tables.Length; i++)
            {
                var styleId = (string?)tables[i].Element(W.tblPr)?.Element(W.tblStyle)?.Attribute(W.val);
                if (styleId is null || !used.Contains(styleId))
                    continue;
                var table = copies[tables[i]];
                var oldMargins = oldTables[i].Element(W.tblPr)?.Element(W.tblCellMar);
                var newMargins = newTables[i].Element(W.tblPr)?.Element(W.tblCellMar);
                if (MarginsDiffer(oldMargins, newMargins))
                    ProjectMarginShell(table, W.tblPr, W.tblCellMar, W.tblPrChange, newMargins, state);

                var cells = tables[i].Descendants(W.tc).Where(c => c.Ancestors(W.tbl).First() == tables[i]).ToArray();
                var oldCells = oldTables[i].Descendants(W.tc).Where(c => c.Ancestors(W.tbl).First() == oldTables[i]).ToArray();
                var newCells = newTables[i].Descendants(W.tc).Where(c => c.Ancestors(W.tbl).First() == newTables[i]).ToArray();
                if (cells.Length != oldCells.Length || cells.Length != newCells.Length)
                    return state.RightSource;
                for (int c = 0; c < cells.Length; c++)
                {
                    var oldOwn = oldCells[c].Element(W.tcPr)?.Element(W.tcMar);
                    var newOwn = newCells[c].Element(W.tcPr)?.Element(W.tcMar);
                    var projected = new XElement(W.tcMar);
                    foreach (var axis in MarginAxes)
                    {
                        var actual = MarginValue(newOwn, axis) ?? MarginValue(newMargins, axis) ?? DefaultMargin(axis);
                        // The table's current margins now match RIGHT, but the retained left style
                        // still supplies old cell/conditional exceptions. Override only a changed
                        // exception; table-wide deltas remain table-wide and leave cell overrides alone.
                        var underOutput = MarginValue(oldOwn, axis) ?? MarginValue(newMargins, axis) ?? DefaultMargin(axis);
                        if (actual != underOutput)
                            projected.Add(MarginElement(newOwn?.Element(axis) ?? newMargins?.Element(axis), axis));
                    }
                    if (projected.HasElements)
                        ProjectMarginShell(copies[cells[c]], W.tcPr, W.tcMar, W.tcPrChange, projected, state);
                }
            }
        }
        return RebindConsumerSources(state.RightSource, sources, copies);
    }

    private static XElement MarginDefinitions(XElement styles) => new("tableMargins",
        styles.Elements(W.style).Where(s => (string?)s.Attribute(W.type) == "table").Select(s =>
            new XElement(W.style, s.Attributes(), s.Element(W.basedOn),
                s.Element(W.tblPr)?.Element(W.tblCellMar), s.Element(W.tcPr)?.Element(W.tcMar),
                s.Elements(W.tblStylePr).Select(region => new XElement(W.tblStylePr, region.Attributes(),
                    region.Element(W.tblPr)?.Element(W.tblCellMar), region.Element(W.tcPr)?.Element(W.tcMar))))));

    private static bool UsesRetainedTableStyle(XElement left, XElement right, string id)
    {
        var seen = new HashSet<string>(StringComparer.Ordinal);
        string? current = id;
        bool retained = false;
        while (current is not null && seen.Count < 64 && seen.Add(current))
        {
            var ls = left.Elements(W.style).Where(s => (string?)s.Attribute(W.type) == "table" && (string?)s.Attribute(W.styleId) == current).ToArray();
            var rs = right.Elements(W.style).Where(s => (string?)s.Attribute(W.type) == "table" && (string?)s.Attribute(W.styleId) == current).ToArray();
            if (ls.Length > 1 || rs.Length != 1 || HasStylePropertyRevisions(rs[0]))
                return false;
            if (ls.Length == 1)
            {
                if (HasStylePropertyRevisions(ls[0]) || !TableChainResolves(left, current))
                    return false;
                retained = true;
            }
            var parent = (string?)rs[0].Element(W.basedOn)?.Attribute(W.val);
            current = parent;
        }
        return current is null && retained;
    }

    private static bool TableChainResolves(XElement styles, string id)
    {
        var seen = new HashSet<string>(StringComparer.Ordinal);
        string? current = id;
        while (current is not null && seen.Count < 64 && seen.Add(current))
        {
            var definitions = styles.Elements(W.style).Where(s => (string?)s.Attribute(W.type) == "table" && (string?)s.Attribute(W.styleId) == current).ToArray();
            if (definitions.Length != 1 || HasStylePropertyRevisions(definitions[0]))
                return false;
            current = (string?)definitions[0].Element(W.basedOn)?.Attribute(W.val);
        }
        return current is null;
    }

    private static string DefaultMargin(XName axis) => axis == W.top || axis == W.bottom ? "0|dxa" : "108|dxa";

    private static string? MarginValue(XElement? margins, XName axis) => margins?.Element(axis) is { } value
        ? ((string?)value.Attribute(W.type) == "nil" ? "0|dxa" : (string?)value.Attribute(W._w) + "|" + (string?)value.Attribute(W.type)) : null;

    private static bool MarginsDiffer(XElement? left, XElement? right) => MarginAxes.Any(axis =>
        (MarginValue(left, axis) ?? DefaultMargin(axis)) != (MarginValue(right, axis) ?? DefaultMargin(axis)));

    private static XElement MarginElement(XElement? value, XName axis) => value is not null
        ? new XElement(value) : new XElement(axis, new XAttribute(W._w, axis == W.top || axis == W.bottom ? 0 : 108), new XAttribute(W.type, "dxa"));

    private static void ProjectMarginShell(XElement host, XName shellName, XName marginsName,
        XName revisionName, XElement? effectiveMargins, RenderState state)
    {
        var original = host.Element(shellName);
        var shell = original is null ? new XElement(shellName) : new XElement(original);
        var margins = shell.Element(marginsName);
        if (margins is null)
        {
            margins = new XElement(marginsName);
            shell.Add(margins);
        }
        // Preserve direct declarations; fill each effective axis independently. This also stops
        // the table normalization pass from substituting border-width padding for style margins.
        foreach (var axis in MarginAxes)
            if (effectiveMargins?.Element(axis) is { } effective)
            {
                margins.Elements(axis).Remove();
                margins.Add(new XElement(effective));
            }
        // Removed style declarations need explicit built-in values under the retained left graph.
        if (shellName == W.tblPr)
            foreach (var axis in new[] { W.top, W.left, W.bottom, W.right })
                if (margins.Element(axis) is null &&
                    !(axis == W.left && margins.Element(W.start) is not null) &&
                    !(axis == W.right && margins.Element(W.end) is not null))
                    margins.Add(MarginElement(null, axis));
        var archive = original is null ? new XElement(shellName) : new XElement(original);
        archive.Elements(revisionName).Remove();
        shell.Elements(revisionName).Remove();
        shell.Add(ProjectionRevision(revisionName, archive, state));
        shell = (XElement)WordprocessingMLUtil.WmlOrderElementsPerStandard(shell);
        original?.Remove();
        InsertShellInSchemaOrder(host, shell, shellName);
    }
}
