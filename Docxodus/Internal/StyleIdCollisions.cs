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
/// out when nothing the output keeps uses it, which is what Word's compare does. When something does — an inserted
/// run's <c>w:rStyle</c>, another copied style's <c>w:basedOn</c> — it is copied under a fresh id and those
/// references follow it.</para>
/// <para>Which side a reference means is read off the reference itself: a <c>w:rStyle</c> names a character style,
/// a <c>w:pStyle</c> a paragraph style, and so on (<see cref="ReferencedType"/>). The original holds no style of the
/// revised style's type under the colliding id, so no valid reference cloned from the original can mean the revised
/// style; every reference of that type to that id came from the revised document, wherever it now sits — inserted
/// content, an equal block emitted from the revised side, or a copied style.</para>
/// </summary>
internal static class StyleIdCollisions
{
    /// <summary>True when <paramref name="stylesRoot"/> already holds a style with <paramref name="styleId"/>
    /// whose type is not <paramref name="type"/>.</summary>
    internal static bool IdHeldByAnotherType(XElement stylesRoot, string? type, string? styleId) =>
        styleId is not null && stylesRoot.Elements(W.style).Any(style =>
            (string?)style.Attribute(W.styleId) == styleId && TypeOf(style) != (type ?? "paragraph"));

    /// <summary>True when any element under <paramref name="roots"/> other than <paramref name="self"/> references a
    /// style of <paramref name="type"/> with <paramref name="styleId"/>.</summary>
    internal static bool IsReferenced(IEnumerable<XElement> roots, string type, string styleId, XElement? self = null) =>
        roots.Any(root => References(root, type, styleId)
            .Any(reference => self is null || !reference.Ancestors().Contains(self)));

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

    private static IEnumerable<XElement> References(XElement root, string type, string styleId) =>
        root.Descendants().Where(e =>
            (string?)e.Attribute(W.val) == styleId && ReferencedType(e) == type);

    private static string TypeOf(XElement style) => (string?)style.Attribute(W.type) ?? "paragraph";
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

    internal bool TryGetFreshId(string type, string styleId, out string freshId) =>
        _fresh.TryGetValue((type, styleId), out freshId!);

    /// <summary>Rewrites every style reference under <paramref name="root"/> whose type and id were remapped.
    /// Returns whether anything changed.</summary>
    internal bool Apply(XElement root)
    {
        if (IsEmpty)
            return false;
        var changed = false;
        foreach (var reference in root.Descendants())
        {
            if ((string?)reference.Attribute(W.val) is not { } id ||
                StyleIdCollisions.ReferencedType(reference) is not { } type ||
                !_fresh.TryGetValue((type, id), out var freshId))
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
}
