// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus.Internal;

/// <summary>
/// Makes the identifiers of drawing objects unique in a package whose story parts hold content from
/// more than one source document (issue #860).
/// <para>A comparison keeps the original's drawing in <c>w:del</c> and the revised document's in
/// <c>w:ins</c>. Both documents number their drawings from 1, so the two <c>wp:docPr/@id</c> values
/// collide. Two VML text boxes likewise bring two copies of the <c>v:shapetype</c> they share, and
/// usually the same shape id (<c>_x0000_s1026</c>), into one part.
/// <see cref="MakeUnique(WordprocessingDocument)"/> renumbers the ids and leaves one shape type
/// definition that every shape can reach.</para>
/// </summary>
internal static class DrawingIds
{
    private static readonly XName W14AnchorId = W14.w14 + "anchorId";
    private static readonly Regex NumberedShapeId = new(@"^_x0000_s(\d{1,9})$", RegexOptions.CultureInvariant);

    /// <summary>
    /// Give every drawing in the story parts of <paramref name="document"/> (body, headers, footers,
    /// footnotes, endnotes, comments) its own <c>wp:docPr/@id</c>; in each part, leave one
    /// <c>v:shapetype</c> per id and give every other VML element its own <c>id</c>. Writes back each part
    /// it changed.
    /// </summary>
    internal static void MakeUnique(WordprocessingDocument document)
    {
        var parts = OwnedPartRelationships.StoryParts(document)
            .Select(owner => (owner.Part, Root: owner.Part.GetXDocument().Root))
            .Where(part => part.Root is not null)
            .ToList();

        var changed = RenumberDocPrIds(parts.Select(part => part.Root!).ToList());
        foreach (var (part, root) in parts)
            if (ResolveShapetypes(root!) | RenumberVmlIds(root!) | changed.Contains(root!))
                part.PutXDocument();
    }

    /// <summary>
    /// Give each drawing in <paramref name="roots"/> its own <c>wp:docPr/@id</c>. The first drawing to use
    /// an id keeps it; a later one gets a fresh id above every id in use. Returns the roots it changed.
    /// </summary>
    internal static HashSet<XElement> RenumberDocPrIds(IReadOnlyList<XElement> roots)
    {
        var drawings = roots
            .SelectMany(root => root.Descendants(WP.docPr))
            .Select(docPr => (Element: docPr, Id: ParseId(docPr)))
            .Where(d => d.Id is not null)
            .Select(d => (d.Element, d.Id!.Value))
            .ToList();

        var inUse = drawings.Select(d => d.Value).ToHashSet();
        var next = inUse.Count == 0 ? 0u : inUse.Max();
        uint Fresh(uint _)
        {
            if (next < uint.MaxValue)
            {
                inUse.Add(++next);
                return next;
            }
            for (uint id = 1; id < uint.MaxValue; id++)
                if (inUse.Add(id))
                    return id;
            throw new InvalidOperationException("no wp:docPr id remains");
        }

        var changed = new HashSet<XElement>();
        foreach (var (docPr, _, assigned) in Deduplicate(drawings, EqualityComparer<uint>.Default, Fresh))
        {
            docPr.SetAttributeValue(NoNamespace.id, assigned.ToString(CultureInfo.InvariantCulture));
            changed.Add(docPr.AncestorsAndSelf().Last());
        }
        return changed;
    }

    /// <summary>
    /// Give every VML element in <paramref name="root"/> other than a <c>v:shapetype</c> its own
    /// <c>id</c>: the first keeps it, a later duplicate gets a fresh one (<c>_x0000_s1027</c> above the
    /// part's highest <c>_x0000_s</c> number, otherwise the id with a numeric suffix). An
    /// <c>o:OLEObject/@ShapeID</c> or <c>w:control/@w:shapeid</c> that named a renumbered element — the
    /// nearest element with that id before it, which is the shape it was written with — follows it.
    /// Returns whether <paramref name="root"/> changed.
    /// </summary>
    internal static bool RenumberVmlIds(XElement root)
    {
        var elements = root.Descendants()
            .Where(e => e.Name.Namespace == VML.vml && e.Name != VML.shapetype)
            .Select(e => (Element: e, Id: (string?)e.Attribute(NoNamespace.id)))
            .Where(e => !string.IsNullOrEmpty(e.Id))
            .Select(e => (e.Element, e.Id!))
            .ToList();
        var duplicated = elements.GroupBy(e => e.Item2, StringComparer.OrdinalIgnoreCase)
            .Where(g => g.Count() > 1)
            .Select(g => g.Key)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        if (duplicated.Count == 0)
            return false;

        var used = root.Descendants()
            .SelectMany(e => new[] { (string?)e.Attribute(NoNamespace.id), (string?)e.Attribute(O.spid) })
            .OfType<string>()
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        var nextNumber = used.Select(id => NumberedShapeId.Match(id))
            .Where(m => m.Success)
            .Select(m => long.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture))
            .DefaultIfEmpty(0)
            .Max();
        string Fresh(string id)
        {
            if (NumberedShapeId.IsMatch(id))
            {
                string candidate;
                do candidate = string.Create(CultureInfo.InvariantCulture, $"_x0000_s{++nextNumber}");
                while (!used.Add(candidate));
                return candidate;
            }
            return FreshSuffixedId(id, used);
        }

        // Bind each reference to the element it names before anything is renamed.
        var referrers = new List<(XElement Referrer, XName Attribute, XElement Target)>();
        var current = new Dictionary<string, XElement>(StringComparer.OrdinalIgnoreCase);
        foreach (var element in root.Descendants())
        {
            if (element.Name.Namespace == VML.vml && element.Name != VML.shapetype &&
                (string?)element.Attribute(NoNamespace.id) is { } id && duplicated.Contains(id))
                current[id] = element;
            foreach (var attribute in new[] { NoNamespace.ShapeID, W.shapeid })
                if ((string?)element.Attribute(attribute) is { } named && current.TryGetValue(named, out var target) &&
                    (element.Name == O.OLEObject || element.Name == W.control))
                    referrers.Add((element, attribute, target));
        }

        var renamed = new Dictionary<XElement, string>();
        foreach (var (element, _, assigned) in Deduplicate(elements.Where(e => duplicated.Contains(e.Item2)),
                     StringComparer.OrdinalIgnoreCase, Fresh))
        {
            element.SetAttributeValue(NoNamespace.id, assigned);
            renamed[element] = assigned;
        }
        foreach (var (referrer, attribute, target) in referrers)
            if (renamed.TryGetValue(target, out var id))
                referrer.SetAttributeValue(attribute, id);
        return renamed.Count > 0;
    }

    /// <summary>
    /// Leave, for each <c>v:shapetype</c> id in <paramref name="root"/>, definitions that every shape can
    /// reach in every view of the document — as it stands, with every revision accepted, and with every
    /// revision rejected — and no two with one id.
    /// <para>When all definitions of an id are equivalent, one definition that no revision or Markup
    /// Compatibility branch can remove is kept and the others are dropped: the first one, if it already
    /// sits outside every revision, branch and text box, otherwise a copy placed in a new untracked run at
    /// the start of the first paragraph that is. Otherwise, or when no paragraph qualifies, each shape keeps
    /// the nearest definition before it that is present in every view the shape is in, and each later
    /// definition that a kept one cannot stand in for gets a fresh id.</para>
    /// Returns whether <paramref name="root"/> changed.
    /// </summary>
    internal static bool ResolveShapetypes(XElement root)
    {
        var definitionsById = root.Descendants(VML.shapetype)
            .GroupBy(definition => (string?)definition.Attribute(NoNamespace.id), StringComparer.Ordinal)
            .Where(group => group.Key is { Length: > 0 })
            .ToDictionary(group => group.Key!, group => group.ToList(), StringComparer.Ordinal);
        if (definitionsById.Count == 0)
            return false;

        var shapesById = root.Descendants()
            .Where(e => e.Name.Namespace == VML.vml && e.Name != VML.shapetype)
            .Select(e => (Shape: e, Type: (string?)e.Attribute(NoNamespace.type)))
            .Where(s => s.Type is { Length: > 1 } && s.Type[0] == '#' && definitionsById.ContainsKey(s.Type[1..]))
            .ToLookup(s => s.Type![1..], s => s.Shape, StringComparer.Ordinal);

        var changed = false;
        HashSet<string>? usedIds = null;
        foreach (var (id, definitions) in definitionsById)
        {
            var shapes = shapesById[id].ToList();
            if (definitions.Count == 1 && !shapes.Any(shape => CanLose(shape, definitions)))
                continue;

            if (definitions.Skip(1).All(d => Equivalent(definitions[0], d)) && Anchor(root, definitions, shapes) is { } kept)
            {
                foreach (var definition in definitions.Where(d => d != kept))
                    definition.Remove();
                changed = true;
                continue;
            }

            if (definitions.Count == 1)
                continue;
            usedIds ??= root.Descendants()
                .Where(e => e.Name.Namespace == VML.vml)
                .Select(e => (string?)e.Attribute(NoNamespace.id))
                .OfType<string>()
                .ToHashSet(StringComparer.Ordinal);
            Rename(definitions, shapes, usedIds, id);
            changed = true;
        }
        return changed;
    }

    /// <summary>The definition to keep for equivalent <paramref name="definitions"/>: the first, if no
    /// revision or branch can remove it, otherwise a new copy in an untracked run, or null when no paragraph
    /// in <paramref name="root"/> can hold one.</summary>
    private static XElement? Anchor(XElement root, List<XElement> definitions, List<XElement> shapes)
    {
        var first = definitions[0];
        if (IsFixed(first))
            return first;

        var earliest = shapes.Append(first).OrderBy(e => e, DocumentOrder.Instance).First();
        var host = earliest.Ancestors(W.p).FirstOrDefault(IsFixed) ?? root.Descendants(W.p).FirstOrDefault(IsFixed);
        if (host is null)
            return null;

        var copy = new XElement(first);
        copy.Attribute(W14AnchorId)?.Remove();
        var picture = new XElement(W.pict, copy);
        foreach (var ns in copy.DescendantsAndSelf()
                     .SelectMany(e => e.Attributes().Where(a => !a.IsNamespaceDeclaration).Select(a => a.Name.Namespace).Append(e.Name.Namespace))
                     .Where(ns => ns != XNamespace.None && host.GetPrefixOfNamespace(ns) is null)
                     .Distinct())
        {
            if (first.GetPrefixOfNamespace(ns) is { } prefix && picture.Attribute(XNamespace.Xmlns + prefix) is null)
                picture.Add(new XAttribute(XNamespace.Xmlns + prefix, ns.NamespaceName));
        }

        var run = new XElement(W.r, picture);
        if (host.Element(W.pPr) is { } properties)
            properties.AddAfterSelf(run);
        else
            host.AddFirst(run);
        return copy;
    }

    /// <summary>Keep the first definition's id; bind each shape to the nearest definition before it that
    /// is present wherever the shape is (or, failing that, after it, or the nearest before it); drop a
    /// later definition an equivalent kept one can stand in for; rename the rest.</summary>
    private static void Rename(List<XElement> definitions, List<XElement> shapes, HashSet<string> usedIds, string id)
    {
        var inOrder = definitions.OrderBy(d => d, DocumentOrder.Instance).ToList();
        var boundTo = new Dictionary<XElement, XElement>();
        foreach (var shape in shapes)
        {
            var before = inOrder.Where(d => DocumentOrder.Instance.Compare(d, shape) < 0).Reverse().ToList();
            var after = inOrder.Where(d => DocumentOrder.Instance.Compare(d, shape) > 0).ToList();
            boundTo[shape] = before.FirstOrDefault(d => PresentWherever(shape, d))
                ?? after.FirstOrDefault(d => PresentWherever(shape, d))
                ?? before.FirstOrDefault()
                ?? inOrder[0];
        }

        var standIn = new Dictionary<XElement, XElement>();
        var kept = new List<XElement> { inOrder[0] };
        foreach (var definition in inOrder.Skip(1))
        {
            if (kept.FirstOrDefault(k => Equivalent(k, definition) && PresentWherever(definition, k)) is { } target)
            {
                standIn[definition] = target;
                definition.Remove();
                continue;
            }
            definition.SetAttributeValue(NoNamespace.id, FreshSuffixedId(id, usedIds));
            kept.Add(definition);
        }

        foreach (var (shape, definition) in boundTo)
        {
            var target = standIn.TryGetValue(definition, out var replacement) ? replacement : definition;
            shape.SetAttributeValue(NoNamespace.type, "#" + (string)target.Attribute(NoNamespace.id)!);
        }
    }

    /// <summary>
    /// Run a first-keeps-it pass over <paramref name="items"/> in document order: an item whose id an
    /// earlier item already holds gets <paramref name="fresh"/>, except that two items offered as
    /// alternatives — their nearest common ancestor is an <c>mc:AlternateContent</c>, so a consumer reads
    /// only one, as Word writes the <c>mc:Choice</c> and <c>mc:Fallback</c> copies of one drawing — are
    /// one object and share an id. Returns the items whose id changed.
    /// </summary>
    private static List<(XElement Element, T Old, T New)> Deduplicate<T>(
        IEnumerable<(XElement Element, T Id)> items, IEqualityComparer<T> comparer, Func<T, T> fresh)
        where T : notnull
    {
        var changes = new List<(XElement Element, T Old, T New)>();
        var taken = new HashSet<T>(comparer);
        var seen = new Dictionary<T, List<(XElement Element, T Assigned)>>(comparer);
        foreach (var (element, id) in items)
        {
            if (!seen.TryGetValue(id, out var earlier))
                seen[id] = earlier = new List<(XElement Element, T Assigned)>();

            var alternative = earlier.FirstOrDefault(e => AreAlternatives(e.Element, element));
            var assigned = alternative.Element is not null ? alternative.Assigned
                : taken.Contains(id) ? fresh(id) : id;
            taken.Add(assigned);
            earlier.Add((element, assigned));
            if (!EqualityComparer<T>.Default.Equals(assigned, id))
                changes.Add((element, id, assigned));
        }
        return changes;
    }

    private static bool AreAlternatives(XElement a, XElement b)
    {
        var ancestorsOfA = a.Ancestors().ToHashSet();
        return b.Ancestors().FirstOrDefault(ancestorsOfA.Contains)?.Name == MC.AlternateContent;
    }

    /// <summary>Whether <paramref name="shape"/> is present, after accepting or after rejecting every
    /// revision, where none of <paramref name="definitions"/> is.</summary>
    private static bool CanLose(XElement shape, List<XElement> definitions) =>
        new[] { true, false }.Any(accept =>
            SurvivesRevisions(shape, accept) && !definitions.Any(d => SurvivesRevisions(d, accept)));

    /// <summary>Whether no revision, Markup Compatibility branch, text box or tracked table row can remove
    /// <paramref name="element"/>.</summary>
    private static bool IsFixed(XElement element) =>
        !element.AncestorsAndSelf().Any(a =>
            a.Name == W.ins || a.Name == W.del || a.Name == W.moveFrom || a.Name == W.moveTo ||
            a.Name == MC.Choice || a.Name == MC.Fallback || a.Name == W.txbxContent ||
            (a.Name == W.tr && a.Element(W.trPr) is { } row && (row.Element(W.ins) is not null || row.Element(W.del) is not null)));

    /// <summary>Whether <paramref name="kept"/> is present in every view in which <paramref name="other"/>
    /// is: after accepting all revisions, after rejecting them all, and under every choice of Markup
    /// Compatibility branches a consumer makes by their <c>Requires</c>.</summary>
    private static bool PresentWherever(XElement other, XElement kept)
    {
        foreach (var accept in new[] { true, false })
            if (SurvivesRevisions(other, accept) && !SurvivesRevisions(kept, accept))
                return false;
        return BranchesOf(kept).IsSubsetOf(BranchesOf(other));
    }

    private static bool SurvivesRevisions(XElement element, bool accept) =>
        !element.Ancestors().Any(a => accept
            ? a.Name == W.del || a.Name == W.moveFrom || IsTrackedRow(a, W.del)
            : a.Name == W.ins || a.Name == W.moveTo || IsTrackedRow(a, W.ins));

    private static bool IsTrackedRow(XElement element, XName mark) =>
        element.Name == W.tr && element.Element(W.trPr)?.Element(mark) is not null;

    /// <summary>The Markup Compatibility branches <paramref name="element"/> lies in, each named by what
    /// makes a consumer choose it: a <c>mc:Choice</c> by its <c>Requires</c>, a <c>mc:Fallback</c> by the
    /// <c>Requires</c> of the choices it stands in for.</summary>
    private static HashSet<string> BranchesOf(XElement element) =>
        element.Ancestors()
            .Where(a => a.Name == MC.Choice || a.Name == MC.Fallback)
            .Select(a => a.Name == MC.Choice
                ? "choice:" + (string?)a.Attribute(NoNamespace.Requires)
                : "fallback:" + string.Join("|", a.Parent!.Elements(MC.Choice).Select(c => (string?)c.Attribute(NoNamespace.Requires))))
            .ToHashSet(StringComparer.Ordinal);

    /// <summary>Whether two definitions describe the same shape type: the same elements and attribute
    /// values throughout, apart from namespace declarations and <c>w14:anchorId</c>, which identifies
    /// the definition rather than describing it.</summary>
    private static bool Equivalent(XElement a, XElement b)
    {
        static HashSet<(XName, string)> Attributes(XElement e) =>
            e.Attributes()
                .Where(attribute => !attribute.IsNamespaceDeclaration && attribute.Name != W14AnchorId)
                .Select(attribute => (attribute.Name, attribute.Value))
                .ToHashSet();
        static string Text(XElement e) => string.Concat(e.Nodes().OfType<XText>().Select(t => t.Value)).Trim();

        return a.Name == b.Name &&
            Attributes(a).SetEquals(Attributes(b)) &&
            Text(a) == Text(b) &&
            a.Elements().Count() == b.Elements().Count() &&
            a.Elements().Zip(b.Elements()).All(pair => Equivalent(pair.First, pair.Second));
    }

    private static uint? ParseId(XElement docPr) =>
        uint.TryParse((string?)docPr.Attribute(NoNamespace.id), NumberStyles.None, CultureInfo.InvariantCulture, out var id)
            ? id
            : null;

    private static string FreshSuffixedId(string id, HashSet<string> usedIds)
    {
        for (var n = 1; ; n++)
        {
            var candidate = string.Create(CultureInfo.InvariantCulture, $"{id}_{n}");
            if (usedIds.Add(candidate))
                return candidate;
        }
    }

    private sealed class DocumentOrder : IComparer<XElement>
    {
        internal static readonly DocumentOrder Instance = new();

        public int Compare(XElement? x, XElement? y) =>
            ReferenceEquals(x, y) ? 0 : XNode.CompareDocumentOrder(x, y);
    }
}
