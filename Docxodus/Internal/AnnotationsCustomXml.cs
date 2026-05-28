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
/// Read/write helpers for the Docxodus annotations Custom XML Part. Shared
/// between the public <see cref="AnnotationManager"/> (WmlDocument API) and
/// <see cref="AnnotationOps"/> (live DocxSession API) so the serialization
/// shape and part-discovery rules live in exactly one place.
/// </summary>
internal static class AnnotationsCustomXml
{
    public const string Namespace = AnnotationManager.AnnotationsNamespace;
    private static readonly XNamespace Ann = Namespace;

    public static CustomXmlPart? Find(WordprocessingDocument doc)
    {
        var main = doc.MainDocumentPart;
        if (main is null) return null;
        foreach (var part in main.CustomXmlParts)
        {
            try
            {
                var xdoc = part.GetXDocument();
                if (xdoc.Root?.Name.Namespace == Ann && xdoc.Root.Name.LocalName == "annotations")
                    return part;
            }
            catch
            {
                // Not XML, or not annotations — skip.
            }
        }
        return null;
    }

    public static CustomXmlPart GetOrCreate(WordprocessingDocument doc)
    {
        var existing = Find(doc);
        if (existing is not null) return existing;

        var part = doc.MainDocumentPart!.AddCustomXmlPart(CustomXmlPartType.CustomXml);
        var xdoc = new XDocument(
            new XDeclaration("1.0", "UTF-8", "yes"),
            new XElement(Ann + "annotations", new XAttribute("version", "1.0")));
        part.PutXDocument(xdoc);
        return part;
    }

    public static DocumentAnnotation? FindById(WordprocessingDocument doc, string annotationId)
    {
        var part = Find(doc);
        if (part is null) return null;
        var element = part.GetXDocument().Root?
            .Elements(Ann + "annotation")
            .FirstOrDefault(a => (string?)a.Attribute("id") == annotationId);
        return element is null ? null : Parse(element);
    }

    public static IReadOnlyList<DocumentAnnotation> ReadAll(WordprocessingDocument doc)
    {
        var part = Find(doc);
        if (part is null) return Array.Empty<DocumentAnnotation>();
        return part.GetXDocument().Root?
            .Elements(Ann + "annotation")
            .Select(Parse)
            .Where(a => a is not null)
            .Select(a => a!)
            .ToList() ?? new List<DocumentAnnotation>();
    }

    public static void Write(WordprocessingDocument doc, DocumentAnnotation annotation)
    {
        var part = GetOrCreate(doc);
        var xdoc = part.GetXDocument();
        var existing = xdoc.Root?
            .Elements(Ann + "annotation")
            .FirstOrDefault(a => (string?)a.Attribute("id") == annotation.Id);
        existing?.Remove();
        xdoc.Root!.Add(Serialize(annotation));
        part.PutXDocument();
    }

    public static bool Remove(WordprocessingDocument doc, string annotationId)
    {
        var part = Find(doc);
        if (part is null) return false;
        var xdoc = part.GetXDocument();
        var element = xdoc.Root?
            .Elements(Ann + "annotation")
            .FirstOrDefault(a => (string?)a.Attribute("id") == annotationId);
        if (element is null) return false;
        element.Remove();
        part.PutXDocument();
        return true;
    }

    private static XElement Serialize(DocumentAnnotation a)
    {
        var element = new XElement(Ann + "annotation",
            new XAttribute("id", a.Id),
            new XAttribute("labelId", a.LabelId ?? ""),
            new XAttribute("label", a.Label ?? ""),
            new XAttribute("color", a.Color ?? "#FFFF00"));
        if (!string.IsNullOrEmpty(a.Author)) element.Add(new XAttribute("author", a.Author));
        if (a.Created.HasValue) element.Add(new XAttribute("created", a.Created.Value.ToString("o")));
        element.Add(new XElement(Ann + "range", new XAttribute("bookmarkName", a.BookmarkName ?? "")));
        if (a.StartPage.HasValue && a.EndPage.HasValue)
        {
            element.Add(new XElement(Ann + "pageSpan",
                new XAttribute("startPage", a.StartPage.Value),
                new XAttribute("endPage", a.EndPage.Value),
                new XAttribute("stale", a.PageInfoStale ? "true" : "false")));
        }
        if (a.Metadata is { Count: > 0 })
        {
            var meta = new XElement(Ann + "metadata");
            foreach (var (key, value) in a.Metadata)
            {
                meta.Add(new XElement(Ann + "item",
                    new XAttribute("key", key),
                    value ?? ""));
            }
            element.Add(meta);
        }
        return element;
    }

    private static DocumentAnnotation? Parse(XElement element)
    {
        var id = (string?)element.Attribute("id");
        if (string.IsNullOrEmpty(id)) return null;

        var a = new DocumentAnnotation
        {
            Id = id,
            LabelId = (string?)element.Attribute("labelId") ?? "",
            Label = (string?)element.Attribute("label") ?? "",
            Color = (string?)element.Attribute("color") ?? "",
            Author = (string?)element.Attribute("author"),
        };

        var createdStr = (string?)element.Attribute("created");
        if (DateTime.TryParse(createdStr, out var created)) a.Created = created;

        var range = element.Element(Ann + "range");
        if (range is not null) a.BookmarkName = (string?)range.Attribute("bookmarkName");

        var pageSpan = element.Element(Ann + "pageSpan");
        if (pageSpan is not null)
        {
            if (int.TryParse((string?)pageSpan.Attribute("startPage"), out var sp)) a.StartPage = sp;
            if (int.TryParse((string?)pageSpan.Attribute("endPage"), out var ep)) a.EndPage = ep;
            a.PageInfoStale = ((string?)pageSpan.Attribute("stale"))?.ToLowerInvariant() == "true";
            if (DateTime.TryParse((string?)pageSpan.Attribute("computedAt"), out var ca))
                a.PageInfoComputedAt = ca;
        }

        var meta = element.Element(Ann + "metadata");
        if (meta is not null)
        {
            foreach (var item in meta.Elements(Ann + "item"))
            {
                var key = (string?)item.Attribute("key");
                if (!string.IsNullOrEmpty(key))
                    a.Metadata[key] = item.Value;
            }
        }
        return a;
    }
}
