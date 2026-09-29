// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus.Internal;

/// <summary>
/// The namespace declarations on the part roots of a set of source documents — the prefix each binds to
/// a namespace, and the namespaces each lists in <c>mc:Ignorable</c> — and the pass that restores them on
/// parts that received content cloned out of those documents.
/// <para>A LINQ to XML clone keeps the names of its elements and attributes but not the declarations of
/// the ancestors it was cloned from. Added to a part whose root lacks them, a Word extension attribute
/// serializes under a prefix the writer invents (<c>p3:durableId</c>) that no <c>mc:Ignorable</c> lists,
/// so a consumer that does not know the namespace must reject it, and an <c>mc:Choice</c> keeps
/// <c>Requires="w14"</c> where nothing declares <c>w14</c> (issue #836). Word writes every declaration on
/// the part root; so does <see cref="DeclareIn(XElement)"/>.</para>
/// </summary>
internal sealed class PartNamespaces
{
    private static readonly char[] ListSeparators = { ' ', '\t', '\r', '\n' };

    private readonly Dictionary<string, XNamespace> _namespaceOfPrefix = new(StringComparer.Ordinal);
    private readonly Dictionary<XNamespace, string> _prefixOfNamespace = new();
    private readonly HashSet<XNamespace> _ignorable = new();

    /// <summary>The declarations on every XML part root of <paramref name="sources"/>. Where two roots
    /// disagree, the first to bind a prefix, or to give a namespace a prefix, wins.</summary>
    internal static PartNamespaces Of(IEnumerable<WmlDocument> sources) =>
        Of(sources.SelectMany(PartRootStartTags).ToArray());

    /// <summary>The declarations on <paramref name="roots"/>; only their own attributes are read.</summary>
    internal static PartNamespaces Of(params XElement[] roots)
    {
        var namespaces = new PartNamespaces();
        foreach (var root in roots)
        {
            foreach (var declaration in root.Attributes().Where(a => a.Name.Namespace == XNamespace.Xmlns))
            {
                namespaces._namespaceOfPrefix.TryAdd(declaration.Name.LocalName, declaration.Value);
                namespaces._prefixOfNamespace.TryAdd(declaration.Value, declaration.Name.LocalName);
            }
            foreach (var prefix in ListedPrefixes(root.Attribute(MC.Ignorable)))
                if (root.GetNamespaceOfPrefix(prefix) is { } ignorable)
                    namespaces._ignorable.Add(ignorable);
        }
        return namespaces;
    }

    /// <summary>Run <see cref="DeclareIn(XElement)"/> on each part of <paramref name="package"/> that has
    /// been loaded as a LINQ to XML tree — the only parts content can have been cloned into — and write
    /// back every part it changed.</summary>
    internal void DeclareIn(OpenXmlPackage package)
    {
        foreach (var part in package.GetAllParts())
            if (part.Annotation<XDocument>()?.Root is { } root && DeclareIn(root))
                part.PutXDocument();
    }

    /// <summary>
    /// Declare on <paramref name="root"/> each namespace its content uses without one in scope — in an
    /// element or attribute name, or as a prefix a Markup Compatibility attribute lists — under the
    /// sources' prefix for it, unless the root already binds that prefix. Then list in
    /// <c>mc:Ignorable</c> each namespace the content uses that a source lists there. Returns whether
    /// <paramref name="root"/> changed.
    /// </summary>
    internal bool DeclareIn(XElement root)
    {
        // XNamespace instances are atomized, so identity is equality — and far cheaper to hash than the URI.
        var boundOnRoot = new HashSet<XNamespace>(
            root.Attributes().Where(a => a.Name.Namespace == XNamespace.Xmlns).Select(a => (XNamespace)a.Value),
            ReferenceEqualityComparer.Instance);
        var used = new List<XNamespace>();
        var unbound = new List<XNamespace>();
        var unresolved = new List<string>();

        void Use(XElement element, XNamespace ns, bool isElementName)
        {
            if (ns == XNamespace.None || ns == XNamespace.Xml)
                return;
            AddOnce(used, ns);
            if (!boundOnRoot.Contains(ns) && !unbound.Contains(ns) && !InScope(element, ns, isElementName))
                unbound.Add(ns);
        }

        foreach (var element in root.DescendantsAndSelf())
        {
            Use(element, element.Name.Namespace, isElementName: true);
            foreach (var attribute in element.Attributes())
            {
                if (attribute.IsNamespaceDeclaration)
                    continue;
                Use(element, attribute.Name.Namespace, isElementName: false);
                if (!ListsPrefixes(element, attribute))
                    continue;
                foreach (var prefix in ListedPrefixes(attribute))
                {
                    if (element.GetNamespaceOfPrefix(prefix) is { } ns)
                        AddOnce(used, ns);
                    else
                        AddOnce(unresolved, prefix);
                }
            }
        }

        var changed = false;
        foreach (var ns in unbound)
            changed |= _prefixOfNamespace.TryGetValue(ns, out var prefix) && Bind(root, prefix, ns);
        foreach (var prefix in unresolved)
        {
            if (!_namespaceOfPrefix.TryGetValue(prefix, out var ns) || !Bind(root, prefix, ns))
                continue;
            AddOnce(used, ns);
            changed = true;
        }
        foreach (var ns in used.Where(_ignorable.Contains))
        {
            if (root.GetPrefixOfNamespace(ns) is not { } prefix ||
                ListedPrefixes(root.Attribute(MC.Ignorable)).Contains(prefix))
                continue;
            EnsureIgnorablePrefix(root, prefix, ns);
            changed = true;
        }
        return changed;
    }

    /// <summary>Declare <paramref name="prefix"/> for <paramref name="ns"/> on <paramref name="root"/> and
    /// list it in the root's <c>mc:Ignorable</c>, keeping the tokens already there.</summary>
    internal static void EnsureIgnorablePrefix(XElement root, string prefix, XNamespace ns)
    {
        if (root.GetNamespaceOfPrefix("mc") != MC.mc)
            root.SetAttributeValue(XNamespace.Xmlns + "mc", MC.mc.NamespaceName);
        if (root.GetNamespaceOfPrefix(prefix) != ns)
            root.SetAttributeValue(XNamespace.Xmlns + prefix, ns.NamespaceName);

        var tokens = ((string?)root.Attribute(MC.Ignorable) ?? string.Empty)
            .Split(ListSeparators, StringSplitOptions.RemoveEmptyEntries)
            .ToList();
        if (!tokens.Contains(prefix, StringComparer.Ordinal))
            tokens.Add(prefix);
        root.SetAttributeValue(MC.Ignorable, string.Join(" ", tokens));
    }

    /// <summary>Bind <paramref name="prefix"/> on <paramref name="root"/> unless the root already binds
    /// it: rebinding would change what the root's own prefix lists refer to.</summary>
    private static bool Bind(XElement root, string prefix, XNamespace ns)
    {
        if (root.GetNamespaceOfPrefix(prefix) is not null)
            return false;
        root.Add(new XAttribute(XNamespace.Xmlns + prefix, ns.NamespaceName));
        return true;
    }

    /// <summary>Whether the writer can name <paramref name="ns"/> at <paramref name="element"/> without
    /// inventing a prefix: an attribute needs a prefixed declaration, an element may use the default one.</summary>
    private static bool InScope(XElement element, XNamespace ns, bool isElementName) =>
        element.GetPrefixOfNamespace(ns) is not null || (isElementName && element.GetDefaultNamespace() == ns);

    /// <summary>Whether <paramref name="attribute"/> is a Markup Compatibility prefix list — any <c>mc:</c>
    /// attribute, or <c>Requires</c> on <c>mc:Choice</c> — whose prefixes must resolve where it appears.</summary>
    private static bool ListsPrefixes(XElement element, XAttribute attribute) =>
        attribute.Name.Namespace == MC.mc || (attribute.Name == NoNamespace.Requires && element.Name == MC.Choice);

    /// <summary>The prefixes a Markup Compatibility list names: <c>w14</c> in both
    /// <c>Ignorable="w14 w15"</c> and <c>ProcessContent="w14:ext"</c>.</summary>
    private static IEnumerable<string> ListedPrefixes(XAttribute? list) =>
        ((string?)list ?? string.Empty)
            .Split(ListSeparators, StringSplitOptions.RemoveEmptyEntries)
            .Select(token => token.Split(':')[0]);

    private static void AddOnce<T>(List<T> list, T item)
    {
        if (!list.Contains(item))
            list.Add(item);
    }

    /// <summary>The root of each XML part in <paramref name="document"/>, as read by
    /// <see cref="StartTag"/>.</summary>
    private static IEnumerable<XElement> PartRootStartTags(WmlDocument document)
    {
        using var package = new ZipArchive(new MemoryStream(document.DocumentByteArray, writable: false));
        foreach (var entry in package.Entries.Where(e => e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)))
            if (StartTag(entry) is { } root)
                yield return root;
    }

    /// <summary>The part's root element, carrying only the declarations and <c>mc:Ignorable</c> of its
    /// start tag — read without parsing the rest of the part, so a large part costs no more than a small
    /// one. A part that cannot be read, or is not well-formed, contributes nothing.</summary>
    private static XElement? StartTag(ZipArchiveEntry entry)
    {
        try
        {
            using var stream = entry.Open();
            using var reader = XmlReader.Create(stream);
            if (reader.MoveToContent() != XmlNodeType.Element)
                return null;
            var root = new XElement(XName.Get(reader.LocalName, reader.NamespaceURI));
            while (reader.MoveToNextAttribute())
            {
                var name = XName.Get(reader.LocalName, reader.NamespaceURI);
                if (reader.Prefix == "xmlns" || name == MC.Ignorable)
                    root.SetAttributeValue(name, reader.Value);
            }
            return root;
        }
        catch (Exception e) when (e is XmlException or ArgumentException or InvalidDataException)
        {
            return null;
        }
    }
}
