// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Buffers.Binary;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace Docxodus.Verification;

/// <summary>
/// XML infoset normalizer used by package-manifest schema v1.
///
/// It ignores only serialization choices: BOM/encoding/XML declaration, namespace prefix and
/// declaration placement, attribute order, quote/entity spelling, empty-element spelling,
/// CDATA-vs-text spelling, and XML's mandated CR/CRLF line-ending normalization.  It deliberately
/// preserves element order, comments, processing instructions, and every text character in opaque
/// Custom/unknown XML.  For known OOXML/OPC parts only, indentation-only text between child
/// elements is ignored unless <c>xml:space="preserve"</c> is in scope.  For OPC metadata, child
/// order is also serialization-only: Default/Override declarations
/// and Relationship elements are sorted by their complete expanded-name/attribute identity.
/// DTDs and external entity resolution are prohibited.
/// </summary>
internal static class XmlSemanticNormalizer
{
    private const string ContentTypesUri = "/[Content_Types].xml";

    internal static XDocument Parse(byte[] bytes, long maxCharacters)
    {
        var settings = new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            IgnoreComments = false,
            IgnoreProcessingInstructions = false,
            IgnoreWhitespace = false,
            MaxCharactersInDocument = maxCharacters,

            // XmlReaderSettings treats 0 as "no limit", so it must carry a real ceiling to be a
            // limit at all. DtdProcessing.Prohibit already makes declared entities unreachable;
            // this keeps the cap correct if that ever relaxes.
            MaxCharactersFromEntities = maxCharacters,
        };
        using var stream = new MemoryStream(bytes, writable: false);
        using var reader = XmlReader.Create(stream, settings);
        return XDocument.Load(reader, LoadOptions.PreserveWhitespace);
    }

    internal static VerificationDigest Digest(
        XDocument document,
        string entryUri,
        bool ignoreFormattingWhitespace)
    {
        using var hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        WriteByte(hash, (byte)'D');
        foreach (var node in document.Nodes())
            WriteNode(hash, node, entryUri, ignoreFormattingWhitespace,
                isDocumentRoot: true, preserveSpace: false);
        return new VerificationDigest
        {
            Algorithm = "SHA-256",
            Value = Convert.ToHexString(hash.GetHashAndReset()).ToLowerInvariant(),
        };
    }

    private static void WriteNode(
        IncrementalHash hash,
        XNode node,
        string entryUri,
        bool ignoreFormattingWhitespace,
        bool isDocumentRoot = false,
        bool preserveSpace = false)
    {
        switch (node)
        {
            case XElement element:
                WriteElement(hash, element, entryUri, ignoreFormattingWhitespace,
                    isDocumentRoot, preserveSpace);
                break;
            case XCData cdata:
                WriteByte(hash, (byte)'T');
                WriteString(hash, cdata.Value);
                break;
            case XText text:
                WriteByte(hash, (byte)'T');
                WriteString(hash, text.Value);
                break;
            case XComment comment:
                WriteByte(hash, (byte)'C');
                WriteString(hash, comment.Value);
                break;
            case XProcessingInstruction instruction:
                WriteByte(hash, (byte)'P');
                WriteString(hash, instruction.Target);
                WriteString(hash, instruction.Data);
                break;
            case XDocumentType documentType:
                WriteByte(hash, (byte)'Y');
                WriteString(hash, documentType.Name);
                WriteString(hash, documentType.PublicId ?? string.Empty);
                WriteString(hash, documentType.SystemId ?? string.Empty);
                WriteString(hash, documentType.InternalSubset ?? string.Empty);
                break;
        }
    }

    private static void WriteElement(
        IncrementalHash hash,
        XElement element,
        string entryUri,
        bool ignoreFormattingWhitespace,
        bool isDocumentRoot,
        bool inheritedPreserveSpace)
    {
        WriteByte(hash, (byte)'E');
        WriteName(hash, element.Name);
        var attributes = element.Attributes()
            .Where(attribute => !attribute.IsNamespaceDeclaration)
            // An element cannot carry two attributes with the same expanded name, so namespace
            // plus local name is already a total order.
            .OrderBy(attribute => attribute.Name.NamespaceName, StringComparer.Ordinal)
            .ThenBy(attribute => attribute.Name.LocalName, StringComparer.Ordinal)
            .ToList();
        WriteInt32(hash, attributes.Count);
        foreach (var attribute in attributes)
        {
            WriteName(hash, attribute.Name);
            WriteString(hash, attribute.Value);
        }

        var space = element.Attribute(XNamespace.Xml + "space")?.Value;
        var preserveSpace = string.Equals(space, "preserve", StringComparison.Ordinal)
            || (inheritedPreserveSpace && !string.Equals(space, "default", StringComparison.Ordinal));
        IEnumerable<XNode> children = CoalesceAdjacentText(element.Nodes());
        if (ignoreFormattingWhitespace && !preserveSpace
            && children.OfType<XElement>().Any())
        {
            children = children.Where(node =>
                node is not XText text || !string.IsNullOrWhiteSpace(text.Value));
        }
        if (isDocumentRoot && IsRelationshipPart(entryUri))
            children = SortOpcMetadataChildren(children, relationshipPart: true);
        else if (isDocumentRoot && string.Equals(entryUri, ContentTypesUri, StringComparison.OrdinalIgnoreCase))
            children = SortOpcMetadataChildren(children, relationshipPart: false);

        var materialized = children.ToList();
        WriteInt32(hash, materialized.Count);
        foreach (var child in materialized)
            WriteNode(hash, child, entryUri, ignoreFormattingWhitespace,
                preserveSpace: preserveSpace);
        WriteByte(hash, (byte)'e');
    }

    private static IEnumerable<XNode> SortOpcMetadataChildren(
        IEnumerable<XNode> children,
        bool relationshipPart)
    {
        var nodes = children.ToList();
        // Whitespace between metadata records is formatting, not opaque application text.  Other
        // node kinds (comments/PIs) remain in their original relative order and therefore remain
        // semantic inputs.
        var elements = nodes.OfType<XElement>().ToList();
        if (nodes.Any(node => node switch
            {
                XElement => false,
                XText text => !string.IsNullOrWhiteSpace(text.Value),
                _ => true,
            }))
            return nodes;

        IEnumerable<XElement> ordered = relationshipPart
            ? elements.OrderBy(ElementAttributeKey("Id"), StringComparer.Ordinal)
                .ThenBy(ElementAttributeKey("Type"), StringComparer.Ordinal)
                .ThenBy(ElementAttributeKey("Target"), StringComparer.Ordinal)
                .ThenBy(ElementAttributeKey("TargetMode"), StringComparer.Ordinal)
            : elements.OrderBy(element => element.Name.LocalName, StringComparer.Ordinal)
                .ThenBy(ElementAttributeKey("PartName"), StringComparer.OrdinalIgnoreCase)
                .ThenBy(ElementAttributeKey("Extension"), StringComparer.OrdinalIgnoreCase)
                .ThenBy(ElementAttributeKey("ContentType"), StringComparer.Ordinal);
        return ordered.Cast<XNode>();
    }

    private static Func<XElement, string> ElementAttributeKey(string localName) =>
        element => element.Attributes().FirstOrDefault(attribute =>
            attribute.Name.LocalName == localName)?.Value ?? string.Empty;

    private static IEnumerable<XNode> CoalesceAdjacentText(IEnumerable<XNode> source)
    {
        StringBuilder? pending = null;
        foreach (var node in source)
        {
            if (node is XText text)
            {
                pending ??= new StringBuilder();
                pending.Append(text.Value);
                continue;
            }
            if (pending is not null)
            {
                yield return new XText(pending.ToString());
                pending = null;
            }
            yield return node;
        }
        if (pending is not null)
            yield return new XText(pending.ToString());
    }

    /// <summary>
    /// Whether a canonical package URI names an OPC relationship part. Shared with
    /// <see cref="PackageManifestGenerator"/> so the normalizer and the generator can never
    /// disagree about which parts receive OPC metadata child ordering.
    /// </summary>
    internal static bool IsRelationshipPart(string uri) =>
        uri.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)
        && uri.Contains("/_rels/", StringComparison.OrdinalIgnoreCase);

    private static void WriteName(IncrementalHash hash, XName name)
    {
        WriteString(hash, name.NamespaceName);
        WriteString(hash, name.LocalName);
    }

    private static void WriteString(IncrementalHash hash, string value)
    {
        var bytes = Encoding.UTF8.GetBytes(value);
        WriteInt32(hash, bytes.Length);
        hash.AppendData(bytes);
    }

    private static void WriteInt32(IncrementalHash hash, int value)
    {
        Span<byte> bytes = stackalloc byte[sizeof(int)];
        BinaryPrimitives.WriteInt32LittleEndian(bytes, value);
        hash.AppendData(bytes);
    }

    private static void WriteByte(IncrementalHash hash, byte value)
    {
        Span<byte> data = stackalloc byte[1];
        data[0] = value;
        hash.AppendData(data);
    }
}
