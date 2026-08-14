// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus.Internal;

/// <summary>
/// Relationship operations whose owner is the package part containing the referring XML.
/// Hyperlinks use this today; image authoring can reuse the same owner lookup/reference-counted
/// cleanup without learning anything about hyperlink markup.
/// </summary>
internal static class OwnedPartRelationships
{
    internal readonly record struct Owner(OpenXmlPart Part, string Scope)
    {
        public string PartUri => Part.Uri.ToString();
    }

    internal static IReadOnlyList<Owner> StoryParts(WordprocessingDocument document)
    {
        var result = new List<Owner>();
        var main = document.MainDocumentPart;
        if (main is null) return result;

        result.Add(new Owner(main, "body"));
        int n = 0;
        foreach (var part in main.HeaderParts) result.Add(new Owner(part, "hdr" + ++n));
        n = 0;
        foreach (var part in main.FooterParts) result.Add(new Owner(part, "ftr" + ++n));
        if (main.FootnotesPart is not null) result.Add(new Owner(main.FootnotesPart, "fn"));
        if (main.EndnotesPart is not null) result.Add(new Owner(main.EndnotesPart, "en"));
        return result;
    }

    internal static Owner? FindOwner(WordprocessingDocument document, XElement element)
    {
        var root = element.AncestorsAndSelf().Last();
        foreach (var owner in StoryParts(document))
            if (ReferenceEquals(owner.Part.GetXDocument().Root, root)) return owner;
        return null;
    }

    internal static IEnumerable<string> ReferencedIds(XElement root, params XName[] attributes) =>
        root.DescendantsAndSelf()
            .SelectMany(e => attributes.Select(a => (string?)e.Attribute(a)))
            .Where(id => !string.IsNullOrEmpty(id))!
            .Cast<string>();

    internal static bool IsReferenced(OpenXmlPart owner, string relationshipId, params XName[] attributes)
    {
        var root = owner.GetXDocument().Root;
        return root is not null && ReferencedIds(root, attributes)
            .Any(id => string.Equals(id, relationshipId, StringComparison.Ordinal));
    }

    internal static bool DeleteReferenceRelationshipIfOrphaned(
        OpenXmlPart owner, string? relationshipId, params XName[] referenceAttributes)
    {
        if (string.IsNullOrEmpty(relationshipId)
            || IsReferenced(owner, relationshipId, referenceAttributes)) return false;
        try
        {
            owner.DeleteReferenceRelationship(relationshipId);
            return true;
        }
        catch (KeyNotFoundException) { return false; }
        catch (ArgumentOutOfRangeException) { return false; }
    }

    /// <summary>
    /// Deletes a child-part relationship only after the owning part's XML has no remaining
    /// references to its id. This is deliberately generic: drawing/image operations can pass
    /// <c>r:embed</c>/<c>r:link</c>; hyperlink code passes <c>r:id</c> to the reference variant.
    /// </summary>
    internal static bool DeletePartRelationshipIfOrphaned(
        OpenXmlPart owner, OpenXmlPart target, params XName[] referenceAttributes)
    {
        var id = owner.GetIdOfPart(target);
        if (IsReferenced(owner, id, referenceAttributes)) return false;
        owner.DeletePart(target);
        return true;
    }

    internal static HyperlinkRelationship FindOrAddExternalHyperlink(OpenXmlPart owner, Uri uri)
    {
        var existing = owner.HyperlinkRelationships.FirstOrDefault(r =>
            r.IsExternal && Uri.Compare(r.Uri, uri, UriComponents.SerializationInfoString,
                UriFormat.UriEscaped, StringComparison.Ordinal) == 0);
        return existing ?? owner.AddHyperlinkRelationship(uri, true);
    }

    /// <summary>Remove every hyperlink relationship no longer referenced by XML in this owner.
    /// Call only after a destructive mutation of that owner; live shared relationships survive
    /// because cleanup is reference-counted against the complete part tree.</summary>
    internal static int SweepOrphanedHyperlinks(OpenXmlPart owner, XName relationshipAttribute)
    {
        int removed = 0;
        foreach (var relationship in owner.HyperlinkRelationships.ToList())
            if (DeleteReferenceRelationshipIfOrphaned(owner, relationship.Id, relationshipAttribute)) removed++;
        return removed;
    }
}
