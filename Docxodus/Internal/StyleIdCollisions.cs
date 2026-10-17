// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus.Internal;

/// <summary>
/// Keeps style ids unique when a comparison copies the revised document's styles into the original's styles part
/// (issue #1184).
/// <para>A <c>w:styleId</c> names one style across all style types, but the copy matched a revised style to an
/// original one by type and id together, so a revised character style whose id the original uses for a paragraph
/// style was copied in beside it under the same id. The original's style keeps the id. The revised style is left
/// out when nothing that enters the output uses it, which is what Word's compare does. When something does — an
/// inserted run's <c>w:rStyle</c>, a copied style's <c>w:basedOn</c> or <c>w:link</c>, a copied numbering level —
/// it is copied under a fresh id and those references follow it.</para>
/// <para>Which of the two styles a reference means is read off the reference itself: a <c>w:rStyle</c> names a
/// character style, a <c>w:pStyle</c> a paragraph style, and so on (<see cref="ReferencedType"/>). A reference of the
/// revised style's type to the colliding id cannot validly mean the original's style. It is rewritten unless it sits
/// in content only the original holds (<see cref="IsOriginalOnly"/>): deleted or moved-away content, a paragraph
/// whose mark is deleted, a wholly deleted table, and the archived old properties of a property change. There the
/// reference is the original document's own (already invalid) one, and rejecting must give it back unchanged.</para>
/// </summary>
internal static class StyleIdCollisions
{
    /// <summary>A style's type; a <c>w:style</c> without <c>w:type</c> is a paragraph style.</summary>
    internal static string TypeOf(XElement style) => NormalizeType((string?)style.Attribute(W.type));

    /// <summary>A <c>w:type</c> value with the paragraph default applied.</summary>
    internal static string NormalizeType(string? type) => type ?? "paragraph";

    /// <summary>Whether <paramref name="style"/> is the style of <paramref name="type"/> (paragraph when null) with
    /// <paramref name="styleId"/>. The one same-type test the style merges share.</summary>
    internal static bool IsStyle(XElement style, string? type, string? styleId) =>
        (string?)style.Attribute(W.styleId) == styleId && TypeOf(style) == NormalizeType(type);

    /// <summary>The style of <paramref name="type"/> with <paramref name="styleId"/> in <paramref name="stylesRoot"/>.</summary>
    internal static XElement? FindStyle(XElement stylesRoot, string? type, string? styleId) =>
        stylesRoot.Elements(W.style).FirstOrDefault(style => IsStyle(style, type, styleId));

    /// <summary>True when <paramref name="stylesRoot"/> already holds a style with <paramref name="styleId"/>
    /// whose type is not <paramref name="type"/>.</summary>
    internal static bool IdHeldByAnotherType(XElement stylesRoot, string? type, string? styleId) =>
        styleId is not null && stylesRoot.Elements(W.style).Any(style =>
            (string?)style.Attribute(W.styleId) == styleId && TypeOf(style) != NormalizeType(type));

    /// <summary>
    /// The styles of <paramref name="sourceStyles"/> whose id <paramref name="destinationStyles"/> holds for a style
    /// of another type and that something entering the output uses. A use counts when it comes from
    /// <paramref name="contentRoots"/> (the stories that carry the source's content), from a source style that is
    /// copied in (one the destination lacks, or a colliding one already found used), or from a source numbering
    /// definition those stories or copied styles reach through <c>w:numId</c>.
    /// </summary>
    internal static HashSet<XElement> UsedCollidingStyles(
        XElement destinationStyles, XElement sourceStyles, XElement? sourceNumbering, IReadOnlyList<XElement> contentRoots)
    {
        var used = new HashSet<XElement>();
        var colliding = new List<XElement>();
        var entering = new List<XElement>();
        foreach (var style in sourceStyles.Elements(W.style))
        {
            var styleId = (string?)style.Attribute(W.styleId);
            if (styleId is null || FindStyle(destinationStyles, TypeOf(style), styleId) is not null)
                continue;
            (IdHeldByAnotherType(destinationStyles, TypeOf(style), styleId) ? colliding : entering).Add(style);
        }
        if (colliding.Count == 0)
            return used;

        // A colliding style found used enters too, and its own references can make another one used.
        bool changed = true;
        while (changed)
        {
            changed = false;
            var copied = entering.Concat(used).ToList();
            var scopes = contentRoots.Concat(copied)
                .Concat(ReachableNumbering(sourceNumbering, contentRoots.Concat(copied))).ToList();
            foreach (var style in colliding)
            {
                if (used.Contains(style) ||
                    !IsReferenced(scopes, TypeOf(style), (string)style.Attribute(W.styleId)!, style))
                    continue;
                used.Add(style);
                changed = true;
            }
        }
        return used;
    }

    /// <summary>The <c>w:num</c> and <c>w:abstractNum</c> definitions of <paramref name="numbering"/> that a
    /// <c>w:numPr/w:numId</c> under <paramref name="roots"/> names.</summary>
    private static IEnumerable<XElement> ReachableNumbering(XElement? numbering, IEnumerable<XElement> roots)
    {
        if (numbering is null)
            yield break;
        var numIds = roots.SelectMany(root => root.Descendants(W.numPr).Elements(W.numId))
            .Select(numId => (string?)numId.Attribute(W.val)).OfType<string>().ToHashSet(StringComparer.Ordinal);
        var abstractIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (var num in numbering.Elements(W.num))
        {
            if (!numIds.Contains((string?)num.Attribute(W.numId) ?? ""))
                continue;
            yield return num;
            if ((string?)num.Element(W.abstractNumId)?.Attribute(W.val) is { } abstractId)
                abstractIds.Add(abstractId);
        }
        foreach (var abstractNum in numbering.Elements(W.abstractNum))
            if (abstractIds.Contains((string?)abstractNum.Attribute(W.abstractNumId) ?? ""))
                yield return abstractNum;
    }

    /// <summary>True when any element under <paramref name="roots"/> outside <paramref name="self"/> and outside
    /// original-only content references a style of <paramref name="type"/> with <paramref name="styleId"/>.</summary>
    private static bool IsReferenced(IEnumerable<XElement> roots, string type, string styleId, XElement self) =>
        roots.Any(root => root.Descendants().Any(reference =>
            (string?)reference.Attribute(W.val) == styleId && ReferencedType(reference) == type &&
            !IsOriginalOnly(reference) && !reference.Ancestors().Contains(self)));

    /// <summary>The first id <paramref name="styleId"/> followed by a number that no style in
    /// <paramref name="stylesRoots"/> uses (compared without case) and <paramref name="taken"/> does not hold.</summary>
    internal static string FreshId(string styleId, IEnumerable<XElement> stylesRoots, ISet<string> taken)
    {
        var used = new HashSet<string>(stylesRoots.SelectMany(root => root.Elements(W.style))
            .Select(style => (string?)style.Attribute(W.styleId) ?? ""), StringComparer.OrdinalIgnoreCase);
        used.UnionWith(taken);
        for (var n = 1; ; n++)
        {
            var candidate = styleId + n.ToString(System.Globalization.CultureInfo.InvariantCulture);
            if (!used.Contains(candidate))
                return candidate;
        }
    }

    /// <summary>The roots of every story a comparison writes content into: the main document, headers, footers,
    /// footnotes, endnotes and comments.</summary>
    internal static IEnumerable<OpenXmlPart> StoryParts(MainDocumentPart main)
    {
        yield return main;
        foreach (var header in main.HeaderParts)
            yield return header;
        foreach (var footer in main.FooterParts)
            yield return footer;
        if (main.FootnotesPart is not null)
            yield return main.FootnotesPart;
        if (main.EndnotesPart is not null)
            yield return main.EndnotesPart;
        if (main.WordprocessingCommentsPart is not null)
            yield return main.WordprocessingCommentsPart;
    }

    /// <summary>The story roots of <paramref name="main"/>.</summary>
    internal static List<XElement> StoryRoots(MainDocumentPart main) =>
        StoryParts(main).Select(part => part.GetXDocument().Root).OfType<XElement>().ToList();

    /// <summary>The type of style <paramref name="reference"/> names, or null when it is not a style reference.
    /// <c>w:basedOn</c> and <c>w:next</c> name a style of their own style's type, and <c>w:link</c> pairs a
    /// paragraph style with a character style.</summary>
    internal static string? ReferencedType(XElement reference)
    {
        if (reference.Name == W.pStyle)
            return "paragraph";
        if (reference.Name == W.rStyle)
            return "character";
        if (reference.Name == W.tblStyle)
            return "table";
        if (reference.Name == W.numStyleLink || reference.Name == W.styleLink)
            return "numbering";
        if (reference.Parent is not { } owner || owner.Name != W.style)
            return null;
        var ownerType = TypeOf(owner);
        if (reference.Name == W.basedOn || reference.Name == W.next)
            return ownerType;
        if (reference.Name == W.link)
            return ownerType == "paragraph" ? "character" : ownerType == "character" ? "paragraph" : null;
        return null;
    }

    /// <summary>
    /// True when <paramref name="element"/> sits in content that only the original document holds in a redline:
    /// inside a <c>w:del</c> or <c>w:moveFrom</c>, inside an archived property change (<c>w:rPrChange</c>,
    /// <c>w:pPrChange</c>, …), in the properties of a paragraph whose mark is deleted or moved away, or in the
    /// properties of a table whose every row is deleted.
    /// </summary>
    internal static bool IsOriginalOnly(XElement element)
    {
        foreach (var ancestor in element.Ancestors())
        {
            if (ancestor.Name == W.del || ancestor.Name == W.moveFrom)
                return true;
            if (ancestor.Name.Namespace == W.w && ancestor.Name.LocalName.EndsWith("Change", StringComparison.Ordinal))
                return true;
            if (ancestor.Name == W.pPr && ancestor.Parent?.Name == W.p &&
                ancestor.Element(W.rPr) is { } markProps &&
                (markProps.Element(W.del) is not null || markProps.Element(W.moveFrom) is not null))
                return true;
            if (ancestor.Name == W.tblPr && ancestor.Parent is { } table && table.Name == W.tbl &&
                table.Elements(W.tr).Any() &&
                table.Elements(W.tr).All(row => row.Elements(W.trPr).Elements(W.del).Any()))
                return true;
        }
        return false;
    }
}

/// <summary>
/// The revised styles a comparison copied in under a fresh id (see <see cref="StyleIdCollisions"/>), and the
/// rewrite of the references that mean them.
/// </summary>
internal sealed class StyleIdRemap
{
    private readonly Dictionary<(string Type, string Id), string> _fresh = new();

    internal bool IsEmpty => _fresh.Count == 0;

    /// <summary>The fresh ids handed out so far, so the next one is not handed out twice.</summary>
    internal ISet<string> FreshIds { get; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

    /// <summary>Records that references to the <paramref name="type"/> style <paramref name="styleId"/> now go to
    /// <paramref name="freshId"/>.</summary>
    internal void Add(string type, string styleId, string freshId)
    {
        _fresh[(type, styleId)] = freshId;
        FreshIds.Add(freshId);
    }

    internal bool Contains(string type, string styleId) => _fresh.ContainsKey((type, styleId));

    /// <summary>Rewrites every style reference under <paramref name="root"/> (itself included) whose type and id
    /// were remapped, except in original-only content. Returns whether anything changed.</summary>
    internal bool Apply(XElement root)
    {
        if (IsEmpty)
            return false;
        var changed = false;
        foreach (var reference in root.DescendantsAndSelf())
        {
            if ((string?)reference.Attribute(W.val) is not { } id ||
                StyleIdCollisions.ReferencedType(reference) is not { } type ||
                !_fresh.TryGetValue((type, id), out var freshId) ||
                StyleIdCollisions.IsOriginalOnly(reference))
                continue;
            reference.SetAttributeValue(W.val, freshId);
            changed = true;
        }
        return changed;
    }

    /// <summary>Applies the remap to each part's root and saves the parts it changed.</summary>
    internal void Apply(IEnumerable<OpenXmlPart?> parts)
    {
        if (IsEmpty)
            return;
        foreach (var part in parts)
        {
            if (part?.GetXDocument().Root is { } root && Apply(root))
                part.PutXDocument();
        }
    }

    /// <summary>The top-level children (numbering definitions, comments) <paramref name="part"/> holds now, so that
    /// <see cref="ApplyToChildrenAddedSince"/> can tell the ones a later merge adds.</summary>
    internal static HashSet<XElement> ChildrenSnapshot(OpenXmlPart? part) =>
        part?.GetXDocument().Root?.Elements().ToHashSet() ?? new HashSet<XElement>();

    /// <summary>Applies the remap to the top-level children of <paramref name="part"/> added since
    /// <paramref name="before"/> was taken: the definitions merged from the revised document. The original's own
    /// definitions keep their references.</summary>
    internal void ApplyToChildrenAddedSince(OpenXmlPart? part, HashSet<XElement> before)
    {
        if (IsEmpty || part?.GetXDocument().Root is not { } root)
            return;
        var changed = false;
        foreach (var child in root.Elements().Where(e => !before.Contains(e)).ToList())
            changed |= Apply(child);
        if (changed)
            part.PutXDocument();
    }
}
