using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus.Ir.Diff;

internal static partial class IrMarkupRenderer
{
    /// <summary>
    /// A defaults-only change cannot be projected into paragraph styles when tables are present:
    /// doing so promotes defaults above conditional table formatting. Resolve each consumer under
    /// both defaults sets, and materialize only changed font/spacing axes on detached right sources.
    /// Native property histories restore the original direct properties under the retained left
    /// defaults. The script and cached IR remain unchanged; normal emit paths use these sources.
    /// </summary>
    private static IrDocument ProjectTableDocDefaults(WmlDocument left, WmlDocument right, RenderState state)
    {
        if (state.Settings.PreserveInputRevisions ||
            !state.Right.Sources.Values.Any(x => x.Descendants(W.tbl).Any()) &&
            !state.Left.Sources.Values.Any(x => x.Descendants(W.tbl).Any()))
            return state.Right;

        using var leftStream = new OpenXmlMemoryStreamDocument(left);
        using var rightStream = new OpenXmlMemoryStreamDocument(right);
        using var leftDoc = leftStream.GetWordprocessingDocument();
        using var rightDoc = rightStream.GetWordprocessingDocument();
        var lm = leftDoc.MainDocumentPart;
        var rm = rightDoc.MainDocumentPart;
        var ls = lm?.StyleDefinitionsPart?.GetXDocument().Root;
        var rs = rm?.StyleDefinitionsPart?.GetXDocument().Root;
        if (lm is null || rm is null || ls is null || rs is null ||
            DocDefaultsPayloadsEqual(ls, rs) || !SameDefinitions(ls, rs) ||
            lm.GlossaryDocumentPart is not null || rm.GlossaryDocumentPart is not null ||
            !PartsEqual(lm.ThemePart, rm.ThemePart) ||
            !PartsEqual(lm.NumberingDefinitionsPart, rm.NumberingDefinitionsPart) ||
            !PartsEqual(lm.DocumentSettingsPart, rm.DocumentSettingsPart) ||
            HasThemeReference(ls) || HasThemeReference(rs) ||
            HasUnsafePresentationConsumer(state.Left, allowTables: true) ||
            HasUnsafePresentationConsumer(state.Right, allowTables: true))
            return state.Right;

        var inherited = AssembleConsumers(right, state.Right, ls.Element(W.docDefaults), replaceDefaults: true);
        var revised = AssembleConsumers(right, state.Right, null, replaceDefaults: false);
        var copies = new Dictionary<XElement, XElement>();
        var sources = new Dictionary<Uri, XDocument>();
        foreach (var (uri, source) in state.Right.Sources)
        {
            var clone = new XDocument(source);
            sources.Add(uri, clone);
            foreach (var pair in source.Descendants().Zip(clone.Descendants()))
                copies.Add(pair.First, pair.Second);
            if (!inherited.TryGetValue(uri, out var oldEffective) || !revised.TryGetValue(uri, out var newEffective))
                continue;
            var originalConsumers = source.Descendants().Where(IsFormattingConsumer).ToArray();
            var oldConsumers = oldEffective.Descendants().Where(IsFormattingConsumer).ToArray();
            var newConsumers = newEffective.Descendants().Where(IsFormattingConsumer).ToArray();
            // The assembler is a presentation resolver, not a structural correspondence oracle.
            // Decline if a shape changed rather than pairing unrelated consumers by ordinal.
            if (originalConsumers.Length != oldConsumers.Length || originalConsumers.Length != newConsumers.Length ||
                !originalConsumers.Select(e => e.Name).SequenceEqual(oldConsumers.Select(e => e.Name)) ||
                !originalConsumers.Select(e => e.Name).SequenceEqual(newConsumers.Select(e => e.Name)))
                return state.Right;
            for (int i = 0; i < originalConsumers.Length; i++)
                ProjectConsumer(copies[originalConsumers[i]], oldConsumers[i], newConsumers[i], state);
        }

        return RebindConsumerSources(state.Right, sources, copies);
    }

    /// <summary>Rebind renderer provenance without changing any content/format fact or mutating a
    /// cached snapshot. Rows and cells need the same detached sources as their enclosing tables.</summary>
    private static IrDocument RebindConsumerSources(IrDocument document,
        Dictionary<Uri, XDocument> sources, Dictionary<XElement, XElement> copies)
    {
        IrProvenance CopySource(IrProvenance source) => source.Element is { } element && copies.TryGetValue(element, out var copy)
            ? new IrProvenance { Element = copy, PartUri = source.PartUri } : source;
        IrBlock CopyBlock(IrBlock block)
        {
            var copied = block with { Source = CopySource(block.Source) };
            if (copied is IrTable table)
                return table with
                {
                    Rows = IrNodeList.From(table.Rows.Select(row => row with
                    {
                        Source = CopySource(row.Source),
                        Cells = IrNodeList.From(row.Cells.Select(cell => cell with
                        {
                            Source = CopySource(cell.Source),
                            Blocks = IrNodeList.From(cell.Blocks.Select(CopyBlock)),
                        })),
                    })),
                };
            return copied;
        }
        return document with
        {
            Sources = sources,
            AnchorIndex = document.AnchorIndex.ToDictionary(entry => entry.Key, entry => CopyBlock(entry.Value)),
        };
    }

    private static bool SameDefinitions(XElement left, XElement right)
    {
        XElement WithoutDefaults(XElement root)
        {
            var clone = new XElement(root);
            clone.Elements(W.docDefaults).Remove();
            StripStyleNoise(clone);
            return clone;
        }
        return XNode.DeepEquals(WithoutDefaults(left), WithoutDefaults(right));
    }

    private static bool IsFormattingConsumer(XElement element) => element.Name == W.p || element.Name == W.r;

    private static Dictionary<Uri, XDocument> AssembleConsumers(
        WmlDocument right, IrDocument ir, XElement? defaults, bool replaceDefaults)
    {
        using var stream = new OpenXmlMemoryStreamDocument(right);
        using var doc = stream.GetWordprocessingDocument();
        foreach (var part in doc.ContentParts())
            if (ir.Sources.TryGetValue(part.Uri, out var source))
                part.GetXDocument().Root!.ReplaceWith(new XElement(source.Root!));
        if (replaceDefaults)
        {
            var styles = doc.MainDocumentPart!.StyleDefinitionsPart!.GetXDocument().Root!;
            styles.Elements(W.docDefaults).Remove();
            if (defaults is not null)
                styles.AddFirst(new XElement(defaults));
        }
        FormattingAssembler.AssembleFormatting(doc, new FormattingAssemblerSettings
        {
            ClearStyles = false,
            RemoveStyleNamesFromParagraphAndRunProperties = false,
            CreateHtmlConverterAnnotationAttributes = false,
        });
        return doc.ContentParts().ToDictionary(part => part.Uri, part => new XDocument(part.GetXDocument()));
    }

    private static void ProjectConsumer(XElement consumer, XElement oldEffective, XElement newEffective, RenderState state)
    {
        bool paragraph = consumer.Name == W.p;
        if (paragraph && !state.Settings.TrackParagraphFormatChanges)
            return;
        var propertyName = paragraph ? W.pPr : W.rPr;
        var original = consumer.Element(propertyName);
        var properties = original is null ? new XElement(propertyName) : new XElement(original);
        var oldProperties = oldEffective.Element(propertyName);
        var newProperties = newEffective.Element(propertyName);
        bool changed = false;
        bool markChanged = false;
        void Axis(XName name, XName attribute, string? fallback)
        {
            string? oldValue = (string?)oldProperties?.Element(name)?.Attribute(attribute) ?? fallback;
            string? newValue = (string?)newProperties?.Element(name)?.Attribute(attribute) ?? fallback;
            if (oldValue == newValue || newValue is null)
                return;
            var child = properties.Element(name);
            if (child is null)
            {
                child = new XElement(name);
                properties.Add(child);
            }
            child.SetAttributeValue(attribute, newValue);
            changed = true;
        }
        if (paragraph)
        {
            Axis(W.spacing, W.before, "0");
            Axis(W.spacing, W.after, "0");
            Axis(W.spacing, W.line, "240");
            Axis(W.spacing, W.lineRule, "auto");
            var mark = new XElement(W.r, original?.Element(W.rPr));
            ProjectConsumer(mark, new XElement(W.r, oldProperties?.Element(W.rPr)),
                new XElement(W.r, newProperties?.Element(W.rPr)), state);
            if (mark.Element(W.rPr)?.Element(W.rPrChange) is not null)
            {
                properties.Elements(W.rPr).Remove();
                properties.Add(new XElement(mark.Element(W.rPr)!));
                markChanged = true;
            }
        }
        else
        {
            Axis(W.rFonts, W.ascii, "Times New Roman");
            Axis(W.rFonts, W.hAnsi, "Times New Roman");
            Axis(W.rFonts, W.eastAsia, "Times New Roman");
            Axis(W.rFonts, W.cs, "Times New Roman");
            Axis(W.sz, W.val, "20");
            Axis(W.szCs, W.val, "20");
        }
        if (!changed && !markChanged)
            return;
        NormalizeStylePropertyOrder(properties, paragraph ? StylePPrChildOrder : StyleRPrChildOrder);
        if (paragraph && changed)
        {
            var archive = new XElement(W.pPr, original?.Attributes(), original?.Elements().Where(e => e.Name != W.rPr && e.Name != W.sectPr && e.Name != W.pPrChange));
            properties.Elements(W.pPrChange).Remove();
            properties.Add(ProjectionRevision(W.pPrChange, archive, state));
        }
        else if (!paragraph)
        {
            properties.Elements(W.rPrChange).Remove();
            properties.Add(ProjectionRevision(W.rPrChange, new XElement(W.rPr, original?.Attributes(), original?.Elements()), state));
        }
        original?.Remove();
        consumer.AddFirst(properties);
    }

    private static XElement ProjectionRevision(XName name, XElement archive, RenderState state)
    {
        var revision = new XElement(name, state.RevisionAttributes(), archive);
        (state.ProjectedPropertyRevisionIds ??= new HashSet<string>(StringComparer.Ordinal))
            .Add((string)revision.Attribute(W.id)!);
        return revision;
    }

    private static void RenumberProjectedProperties(WordprocessingDocument document, RenderState state)
    {
        if (state.ProjectedPropertyRevisionIds is not { } ids)
            return;
        // A source run may be sliced into several emitted runs. Give every copied property
        // revision its own identity after all emission paths have finished.
        foreach (var part in document.ContentParts())
        {
            foreach (var revision in part.GetXDocument().Descendants().Where(e =>
                         (e.Name == W.rPrChange || e.Name == W.pPrChange) &&
                         ids.Contains((string?)e.Attribute(W.id) ?? "")))
                revision.SetAttributeValue(W.id, state.NextId());
            part.PutXDocument();
        }
    }
}
