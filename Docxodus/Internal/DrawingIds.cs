// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus.Internal;

/// <summary>
/// Makes the identifiers of drawing objects unique in a package whose story parts hold content from
/// more than one source document (issue #860).
/// <para>A comparison keeps the original's drawing in <c>w:del</c> and the revised document's in
/// <c>w:ins</c>. Both documents number their drawings from 1, so the two <c>wp:docPr/@id</c> values
/// collide, and two VML text boxes each bring their own copy of the <c>v:shapetype</c> they share an id
/// with. <see cref="MakeUnique(WordprocessingDocument)"/> renumbers the first and resolves the second.</para>
/// </summary>
internal static class DrawingIds
{
    private static readonly XName W14AnchorId = W14.w14 + "anchorId";

    /// <summary>
    /// Give every drawing in the story parts of <paramref name="document"/> (body, headers, footers,
    /// footnotes, endnotes, comments) its own <c>wp:docPr/@id</c>, and every <c>v:shapetype</c> in one
    /// part its own <c>id</c>. Writes back each part it changed.
    /// </summary>
    internal static void MakeUnique(WordprocessingDocument document)
    {
        var parts = OwnedPartRelationships.StoryParts(document)
            .Select(owner => (owner.Part, Root: owner.Part.GetXDocument().Root))
            .Where(part => part.Root is not null)
            .ToList();

        var changed = RenumberDocPrIds(parts.Select(part => part.Root!).ToList());
        foreach (var (part, root) in parts)
            if (ResolveShapetypes(root!) | changed.Contains(root!))
                part.PutXDocument();
    }

    /// <summary>
    /// Give each drawing in <paramref name="roots"/> its own <c>wp:docPr/@id</c>. The first drawing to use
    /// an id keeps it; a later one gets a fresh id above every id in use. The copies of one drawing that
    /// Markup Compatibility offers as alternatives (the <c>mc:Choice</c> and <c>mc:Fallback</c> of one
    /// <c>mc:AlternateContent</c>) are one object, as Word writes them, and share an id. Returns the roots
    /// it changed.
    /// </summary>
    internal static HashSet<XElement> RenumberDocPrIds(IReadOnlyList<XElement> roots)
    {
        var changed = new HashSet<XElement>();
        var drawings = roots
            .SelectMany(root => root.Descendants(WP.docPr).Select(docPr => (Root: root, DocPr: docPr)))
            .Select(d => (d.Root, d.DocPr, Id: ParseId(d.DocPr)))
            .Where(d => d.Id is not null)
            .ToList();

        var inUse = drawings.Select(d => d.Id!.Value).ToHashSet();
        var next = inUse.Count == 0 ? 0u : inUse.Max();
        var taken = new HashSet<uint>();
        var seen = new Dictionary<uint, List<(XElement DocPr, uint Assigned)>>();
        foreach (var (root, docPr, id) in drawings)
        {
            var original = id!.Value;
            if (!seen.TryGetValue(original, out var earlier))
                seen[original] = earlier = new List<(XElement DocPr, uint Assigned)>();

            var assigned = earlier.FirstOrDefault(e => AreAlternatives(e.DocPr, docPr)) is { DocPr: not null } alternative
                ? alternative.Assigned
                : taken.Contains(original) ? Fresh(ref next, inUse) : original;
            taken.Add(assigned);
            earlier.Add((docPr, assigned));
            if (assigned == original)
                continue;
            docPr.SetAttributeValue(NoNamespace.id, assigned.ToString(CultureInfo.InvariantCulture));
            changed.Add(root);
        }
        return changed;
    }

    /// <summary>
    /// Leave one <c>v:shapetype</c> per <c>id</c> in <paramref name="root"/>. A later definition that is
    /// equivalent to a kept one is removed when the kept one is present in every view that shows it —
    /// after accepting all revisions, after rejecting them all, and under each choice of Markup
    /// Compatibility branches — and the shapes that referenced it use the kept one. Any other later definition keeps its content under a fresh id,
    /// and the shapes that referenced it follow it. A shape refers to the nearest definition of its type
    /// before it, where the definition and its shape came in together from one source. Returns whether
    /// <paramref name="root"/> changed.
    /// </summary>
    internal static bool ResolveShapetypes(XElement root)
    {
        var duplicated = root.Descendants(VML.shapetype)
            .GroupBy(definition => (string?)definition.Attribute(NoNamespace.id), StringComparer.Ordinal)
            .Where(group => group.Key is { Length: > 0 } && group.Count() > 1)
            .ToDictionary(group => group.Key!, group => group.ToList(), StringComparer.Ordinal);
        if (duplicated.Count == 0)
            return false;

        var shapesOf = ShapesByDefinition(root, duplicated);
        var usedIds = root.Descendants()
            .Where(element => element.Name.Namespace == VML.vml)
            .Select(element => (string?)element.Attribute(NoNamespace.id))
            .OfType<string>()
            .ToHashSet(StringComparer.Ordinal);

        foreach (var (id, definitions) in duplicated)
        {
            var kept = new List<XElement> { definitions[0] };
            foreach (var definition in definitions.Skip(1))
            {
                var target = kept.FirstOrDefault(k => Equivalent(k, definition) && PresentWherever(definition, k));
                if (target is not null)
                    definition.Remove();
                else
                {
                    target = definition;
                    target.SetAttributeValue(NoNamespace.id, FreshShapetypeId(id, usedIds));
                    kept.Add(target);
                }
                foreach (var shape in shapesOf[definition])
                    shape.SetAttributeValue(NoNamespace.type, "#" + (string)target.Attribute(NoNamespace.id)!);
            }
        }
        return true;
    }

    /// <summary>The shapes that refer to each duplicated definition: a shape whose <c>type</c> names a
    /// duplicated id refers to the last definition of it before the shape, or to the first one if none
    /// precedes it.</summary>
    private static Dictionary<XElement, List<XElement>> ShapesByDefinition(
        XElement root, Dictionary<string, List<XElement>> duplicated)
    {
        var shapesOf = duplicated.Values.SelectMany(d => d).ToDictionary(d => d, _ => new List<XElement>());
        var current = duplicated.ToDictionary(pair => pair.Key, pair => pair.Value[0], StringComparer.Ordinal);
        foreach (var element in root.Descendants().Where(e => e.Name.Namespace == VML.vml))
        {
            if (element.Name == VML.shapetype)
            {
                if (shapesOf.ContainsKey(element))
                    current[(string)element.Attribute(NoNamespace.id)!] = element;
            }
            else if ((string?)element.Attribute(NoNamespace.type) is { } type && type.StartsWith('#') &&
                     current.TryGetValue(type[1..], out var definition))
                shapesOf[definition].Add(element);
        }
        return shapesOf;
    }

    /// <summary>Whether <paramref name="a"/> and <paramref name="b"/> are the same drawing offered as
    /// alternatives: their nearest common ancestor is an <c>mc:AlternateContent</c>, so they sit in
    /// different branches of it and a consumer reads only one.</summary>
    private static bool AreAlternatives(XElement a, XElement b)
    {
        var ancestorsOfA = a.Ancestors().ToHashSet();
        return b.Ancestors().FirstOrDefault(ancestorsOfA.Contains)?.Name == MC.AlternateContent;
    }

    /// <summary>Whether <paramref name="kept"/> is present in every view in which <paramref name="dropped"/>
    /// is: after accepting all revisions, after rejecting them all, and under every choice of Markup
    /// Compatibility branches a consumer makes by their <c>Requires</c>.</summary>
    private static bool PresentWherever(XElement dropped, XElement kept)
    {
        foreach (var accept in new[] { true, false })
            if (SurvivesRevisions(dropped, accept) && !SurvivesRevisions(kept, accept))
                return false;
        return BranchesOf(kept).IsSubsetOf(BranchesOf(dropped));
    }

    private static bool SurvivesRevisions(XElement element, bool accept) =>
        !element.Ancestors().Any(a => accept ? a.Name == W.del || a.Name == W.moveFrom : a.Name == W.ins || a.Name == W.moveTo);

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

    /// <summary>An id above every id in use, or the smallest unused one once the range above is spent.</summary>
    private static uint Fresh(ref uint next, HashSet<uint> inUse)
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

    private static string FreshShapetypeId(string id, HashSet<string> usedIds)
    {
        for (var n = 1; ; n++)
        {
            var candidate = string.Create(CultureInfo.InvariantCulture, $"{id}_{n}");
            if (usedIds.Add(candidate))
                return candidate;
        }
    }
}
