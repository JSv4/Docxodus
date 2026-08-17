// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Buffers.Binary;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace Docxodus.Verification;

internal sealed class XmlDepthLimitException : XmlException
{
    public XmlDepthLimitException(string message) : base(message)
    {
    }
}

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
    // Recursive canonical emission must have a process-safe structural bound independent of the
    // byte/character ceilings. 256 is far above normal OPC/OOXML nesting while keeping hostile
    // documents comfortably below the runtime stack limit.
    internal const int MaxElementDepth = 256;

    private const string ContentTypesUri = "/[Content_Types].xml";
    private const string XmlSchemaInstanceNamespace =
        "http://www.w3.org/2001/XMLSchema-instance";
    private const string MarkupCompatibilityNamespace =
        "http://schemas.openxmlformats.org/markup-compatibility/2006";

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
        var document = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        EnsureBoundedDepth(document);
        return document;
    }

    internal static VerificationDigest Digest(
        XDocument document,
        string entryUri,
        bool ignoreFormattingWhitespace,
        Func<XAttribute, bool>? includeAttribute = null,
        Func<XAttribute, string>? attributeValueNormalizer = null)
    {
        EnsureBoundedDepth(document);
        using var hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        WriteByte(hash, (byte)'D');
        // XML permits only production-S whitespace outside the document element. It is document
        // serialization formatting, just like an XML declaration, and is never application text.
        foreach (var node in document.Nodes().Where(node =>
                     node is not XText text || !IsXmlWhitespace(text.Value)))
            WriteNode(hash, node, entryUri, ignoreFormattingWhitespace,
                includeAttribute, attributeValueNormalizer,
                isDocumentRoot: true, preserveSpace: false);
        return new VerificationDigest
        {
            Algorithm = "SHA-256",
            Value = Convert.ToHexString(hash.GetHashAndReset()).ToLowerInvariant(),
        };
    }

    internal static VerificationDigest Digest(
        XElement element,
        string entryUri,
        bool ignoreFormattingWhitespace,
        Func<XAttribute, bool>? includeAttribute = null,
        Func<XAttribute, string>? attributeValueNormalizer = null)
    {
        ArgumentNullException.ThrowIfNull(element);
        EnsureBoundedDepth(element);
        using var hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        WriteByte(hash, (byte)'D');
        // Hash the attached element directly. Cloning it into a new XDocument discards namespace
        // bindings inherited from ancestors, making equivalent xsi:type/mc prefix spellings differ.
        WriteElement(hash, element, entryUri, ignoreFormattingWhitespace,
            includeAttribute, attributeValueNormalizer,
            isDocumentRoot: true, inheritedPreserveSpace: false);
        return new VerificationDigest
        {
            Algorithm = "SHA-256",
            Value = Convert.ToHexString(hash.GetHashAndReset()).ToLowerInvariant(),
        };
    }

    private static void EnsureBoundedDepth(XContainer container)
    {
        var pending = new Stack<(XElement Element, int Depth)>();
        if (container is XElement element)
            pending.Push((element, 1));
        else
            foreach (var root in container.Elements())
                pending.Push((root, 1));

        while (pending.Count > 0)
        {
            var (current, depth) = pending.Pop();
            if (depth > MaxElementDepth)
                throw new XmlDepthLimitException(
                    $"XML element nesting exceeds the {MaxElementDepth} level safety limit.");
            foreach (var child in current.Elements())
                pending.Push((child, depth + 1));
        }
    }

    private static void WriteNode(
        IncrementalHash hash,
        XNode node,
        string entryUri,
        bool ignoreFormattingWhitespace,
        Func<XAttribute, bool>? includeAttribute,
        Func<XAttribute, string>? attributeValueNormalizer,
        bool isDocumentRoot = false,
        bool preserveSpace = false)
    {
        switch (node)
        {
            case XElement element:
                WriteElement(hash, element, entryUri, ignoreFormattingWhitespace,
                    includeAttribute, attributeValueNormalizer, isDocumentRoot, preserveSpace);
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
        Func<XAttribute, bool>? includeAttribute,
        Func<XAttribute, string>? attributeValueNormalizer,
        bool isDocumentRoot,
        bool inheritedPreserveSpace)
    {
        WriteByte(hash, (byte)'E');
        WriteName(hash, element.Name);
        var attributes = element.Attributes()
            .Where(attribute => !attribute.IsNamespaceDeclaration
                && (includeAttribute?.Invoke(attribute) ?? true))
            // An element cannot carry two attributes with the same expanded name, so namespace
            // plus local name is already a total order.
            .OrderBy(attribute => attribute.Name.NamespaceName, StringComparer.Ordinal)
            .ThenBy(attribute => attribute.Name.LocalName, StringComparer.Ordinal)
            .ToList();
        WriteInt32(hash, attributes.Count);
        foreach (var attribute in attributes)
        {
            WriteName(hash, attribute.Name);
            WriteAttributeValue(hash, element, attribute, attributeValueNormalizer);
        }

        var space = element.Attribute(XNamespace.Xml + "space")?.Value;
        var preserveSpace = string.Equals(space, "preserve", StringComparison.Ordinal)
            || (inheritedPreserveSpace && !string.Equals(space, "default", StringComparison.Ordinal));
        IEnumerable<XNode> children = CoalesceAdjacentText(element.Nodes());
        if (ignoreFormattingWhitespace && !preserveSpace
            && (isDocumentRoot || children.OfType<XElement>().Any()))
        {
            children = children.Where(node =>
                node is not XText text || !IsXmlWhitespace(text.Value));
        }
        if (isDocumentRoot && IsRelationshipPart(entryUri))
            children = SortOpcMetadataChildren(children, relationshipPart: true);
        else if (isDocumentRoot && string.Equals(entryUri, ContentTypesUri, StringComparison.OrdinalIgnoreCase))
            children = SortOpcMetadataChildren(children, relationshipPart: false);

        var materialized = children.ToList();
        WriteInt32(hash, materialized.Count);
        foreach (var child in materialized)
            WriteNode(hash, child, entryUri, ignoreFormattingWhitespace,
                includeAttribute, attributeValueNormalizer,
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
                XText text => !IsXmlWhitespace(text.Value),
                _ => true,
            }))
            return nodes;

        IEnumerable<XElement> ordered = relationshipPart
            ? elements.OrderBy(ElementAttributeKey("Id"), StringComparer.Ordinal)
                .ThenBy(ElementAttributeKey("Type"), StringComparer.Ordinal)
                .ThenBy(ElementAttributeKey("Target"), StringComparer.Ordinal)
                .ThenBy(ElementAttributeKey("TargetMode"), StringComparer.Ordinal)
                .ThenBy(NormalizedElementIdentity, StringComparer.Ordinal)
            : elements.OrderBy(element => element.Name.LocalName, StringComparer.Ordinal)
                .ThenBy(ElementAttributeKey("PartName"), AsciiCaseInsensitiveComparer.Instance)
                .ThenBy(ElementAttributeKey("Extension"), AsciiCaseInsensitiveComparer.Instance)
                .ThenBy(ElementAttributeKey("ContentType"), StringComparer.Ordinal)
                .ThenBy(NormalizedElementIdentity, StringComparer.Ordinal);
        return ordered.Cast<XNode>();
    }

    private static Func<XElement, string> ElementAttributeKey(string localName) =>
        element => element.Attribute(XName.Get(localName))?.Value ?? string.Empty;

    // The schema keys above make manifests easy to inspect, but they are not a total order: OPC
    // part identifiers fold ASCII case, and relationship markup can carry extension attributes.
    // Use the same normalized token stream that is ultimately hashed as the final tie-breaker so
    // child order remains serialization-only even when primary keys compare equal.
    private static string NormalizedElementIdentity(XElement element)
    {
        using var hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        WriteElement(hash, element, entryUri: string.Empty, ignoreFormattingWhitespace: true,
            includeAttribute: null, attributeValueNormalizer: null,
            isDocumentRoot: false, inheritedPreserveSpace: false);
        return Convert.ToHexString(hash.GetHashAndReset()).ToLowerInvariant();
    }

    private static void WriteAttributeValue(
        IncrementalHash hash,
        XElement context,
        XAttribute attribute,
        Func<XAttribute, string>? attributeValueNormalizer)
    {
        var value = attributeValueNormalizer?.Invoke(attribute) ?? attribute.Value;
        if (attribute.Name.NamespaceName == XmlSchemaInstanceNamespace
            && attribute.Name.LocalName == "type"
            && TryResolveQName(context, value, out var typeName))
        {
            WriteByte(hash, (byte)'Q');
            WriteName(hash, typeName);
            return;
        }

        var isMcAttribute = attribute.Name.NamespaceName == MarkupCompatibilityNamespace;
        var isPrefixList = isMcAttribute
            && attribute.Name.LocalName is "Ignorable" or "MustUnderstand"
            || attribute.Name.NamespaceName.Length == 0
            && attribute.Name.LocalName == "Requires"
            && context.Name.NamespaceName == MarkupCompatibilityNamespace;
        if (isPrefixList
            && TryResolvePrefixList(context, value, out var namespaceNames))
        {
            WriteByte(hash, (byte)'P');
            WriteInt32(hash, namespaceNames.Count);
            foreach (var namespaceName in namespaceNames)
                WriteString(hash, namespaceName);
            return;
        }

        if (isMcAttribute
            && attribute.Name.LocalName is
                "PreserveAttributes" or "PreserveElements" or "ProcessContent"
            && TryResolveQNameList(context, value, out var names))
        {
            WriteByte(hash, (byte)'L');
            WriteInt32(hash, names.Count);
            foreach (var name in names)
                WriteName(hash, name);
            return;
        }

        WriteByte(hash, (byte)'V');
        WriteString(hash, value);
    }

    private static bool TryResolveQName(XElement context, string value, out XName name)
    {
        name = XName.Get("invalid");
        var tokens = SplitXmlWhitespace(value);
        if (tokens.Count != 1)
            return false;
        var token = tokens[0];
        var separator = token.IndexOf(':');
        if (separator != token.LastIndexOf(':'))
            return false;
        var prefix = separator < 0 ? string.Empty : token[..separator];
        var localName = separator < 0 ? token : token[(separator + 1)..];
        try
        {
            if (prefix.Length > 0)
                XmlConvert.VerifyNCName(prefix);
            XmlConvert.VerifyNCName(localName);
        }
        catch (XmlException)
        {
            return false;
        }
        var namespaceName = prefix.Length == 0
            ? context.GetDefaultNamespace()
            : context.GetNamespaceOfPrefix(prefix);
        if (namespaceName is null)
            return false;
        name = namespaceName + localName;
        return true;
    }

    private static bool TryResolveQNameList(
        XElement context,
        string value,
        out IReadOnlyList<XName> names)
    {
        var resolved = new List<XName>();
        foreach (var token in SplitXmlWhitespace(value))
        {
            if (!TryResolveQName(context, token, out var name))
            {
                names = Array.Empty<XName>();
                return false;
            }
            resolved.Add(name);
        }
        names = resolved
            .Distinct()
            .OrderBy(name => name.NamespaceName, StringComparer.Ordinal)
            .ThenBy(name => name.LocalName, StringComparer.Ordinal)
            .ToList();
        return names.Count > 0;
    }

    private static bool TryResolvePrefixList(
        XElement context,
        string value,
        out IReadOnlyList<string> namespaceNames)
    {
        var resolved = new List<string>();
        foreach (var prefix in SplitXmlWhitespace(value))
        {
            try
            {
                XmlConvert.VerifyNCName(prefix);
            }
            catch (XmlException)
            {
                namespaceNames = Array.Empty<string>();
                return false;
            }
            var namespaceName = context.GetNamespaceOfPrefix(prefix);
            if (namespaceName is null)
            {
                namespaceNames = Array.Empty<string>();
                return false;
            }
            resolved.Add(namespaceName.NamespaceName);
        }
        namespaceNames = resolved.Distinct(StringComparer.Ordinal)
            .OrderBy(name => name, StringComparer.Ordinal)
            .ToList();
        return namespaceNames.Count > 0;
    }

    private static IReadOnlyList<string> SplitXmlWhitespace(string value) =>
        value.Split([' ', '\t', '\r', '\n'], StringSplitOptions.RemoveEmptyEntries);

    private static bool IsXmlWhitespace(string value) =>
        value.All(character => character is ' ' or '\t' or '\r' or '\n');

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

    private sealed class AsciiCaseInsensitiveComparer : IComparer<string>
    {
        public static readonly AsciiCaseInsensitiveComparer Instance = new();

        public int Compare(string? left, string? right)
        {
            if (ReferenceEquals(left, right))
                return 0;
            if (left is null)
                return -1;
            if (right is null)
                return 1;
            var sharedLength = Math.Min(left.Length, right.Length);
            for (var index = 0; index < sharedLength; index++)
            {
                var leftCharacter = FoldAscii(left[index]);
                var rightCharacter = FoldAscii(right[index]);
                if (leftCharacter != rightCharacter)
                    return leftCharacter.CompareTo(rightCharacter);
            }
            return left.Length.CompareTo(right.Length);
        }

        private static char FoldAscii(char value) =>
            value is >= 'a' and <= 'z' ? (char)(value - ('a' - 'A')) : value;
    }
}
