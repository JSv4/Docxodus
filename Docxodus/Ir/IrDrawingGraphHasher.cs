#nullable enable

using System;
using System.Buffers.Binary;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus.Ir;

/// <summary>
/// Per-graph resource limits. The production defaults are intentionally generous for ordinary chart and
/// SmartArt packages; the injectable value lets focused tests exercise the conservative cutoff path without
/// allocating a multi-megabyte fixture.
/// </summary>
internal readonly record struct IrDrawingGraphHashLimits(
    int MaxXmlGraphDepth,
    int MaxXmlGraphParts,
    int MaxXmlPartBytes,
    long MaxXmlGraphBytes)
{
    public static readonly IrDrawingGraphHashLimits Default = new(
        MaxXmlGraphDepth: 32,
        MaxXmlGraphParts: 256,
        MaxXmlPartBytes: 16 * 1024 * 1024,
        MaxXmlGraphBytes: 64L * 1024 * 1024);
}

/// <summary>
/// Computes a drawing-local identity that includes the semantic content of reachable XML relationship graphs.
/// The ordinary <see cref="IrRelResolver"/> deliberately collapses XML targets to a single token for broad
/// legacy parity; that is too lossy for a <c>w:drawing</c>, whose visible presentation can change solely in a
/// chart, SmartArt, or related XML part while its outer DrawingML remains byte-identical.
/// </summary>
/// <remarks>
/// The scope is intentionally limited to relationship attribute names that the renderer imports. Binary,
/// external, dangling, and unsupported relationship attributes retain <see cref="IrRelResolver"/>'s established
/// behavior at the outer drawing; graph-internal binary/external leaves add package-semantic framing so a nested
/// relationship-type or content-type change is not hidden. This keeps ordinary image identity stable while making
/// opaque XML-backed drawings reversible.
/// </remarks>
internal sealed class IrDrawingGraphHasher
{
    // These bounds apply to one outer drawing hash, not a whole document read. They prevent a malformed package
    // from turning opaque drawing identity into an unbounded graph walk while remaining far above normal chart /
    // SmartArt package sizes.
    private static readonly XNamespace R =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace Dsp =
        "http://schemas.microsoft.com/office/drawing/2008/diagram";
    private static readonly XName DspDataModelExt = Dsp + "dataModelExt";

    // Keep this aligned with ComparisonUnitWord.s_RelationshipAttributeNames, the relationship edge set the
    // renderer imports. Hashing an edge the renderer cannot carry would make an edit detectable but not safely
    // renderable, so unknown r:* attributes intentionally retain the legacy resolver behavior.
    private static readonly HashSet<XName> ImportedRelationshipAttributes = new()
    {
        R + "embed",
        R + "link",
        R + "id",
        R + "cs",
        R + "dm",
        R + "lo",
        R + "qs",
        R + "href",
        R + "pict",
    };

    private readonly OpenXmlPart _drawingOwner;
    // A cutoff cannot prove the remaining graph is equal. This lazy fingerprint binds an incomplete graph to
    // the exact normalized source package, making such comparisons conservative: unrelated package changes can
    // produce an extra replacement, but a changed child beyond a resource boundary cannot leak through Reject.
    private readonly Lazy<IrHash> _sourceDocumentFingerprint;
    private readonly IrDrawingGraphHashLimits _limits;
    private readonly Dictionary<OpenXmlPart, IrRelResolver> _legacyResolvers = new();
    // A cached value is keyed by the remaining traversal depth. A shallow complete walk must not make a later,
    // deep edge bypass its safety cutoff (or vice versa).
    private readonly Dictionary<XmlCacheKey, string> _xmlIdentityCache = new();
    // A resource-cutoff identity includes the full raw bytes of its immediate XML part. Reuse that bounded-memory
    // streaming digest when multiple drawings reference the same oversized part, while keeping incomplete graph
    // identities themselves uncached because their meaning depends on the traversal budget at the call site.
    private readonly Dictionary<OpenXmlPart, RawPartIdentity> _rawPartIdentityCache = new();

    private readonly record struct XmlCacheKey(
        OpenXmlPart Part, OpenXmlPart? DiagramDataRelationshipOwner, int RemainingDepth);
    private readonly record struct GraphIdentity(string Value, bool Cacheable);
    private readonly record struct RawPartIdentity(long Length, byte[] Digest);
    private readonly record struct XmlReadResult(byte[]? Bytes, string? LimitReason)
    {
        public bool LimitReached => LimitReason is not null;
    }

    private sealed class Traversal
    {
        public HashSet<OpenXmlPart> ActiveXmlParts { get; } = new();
        public HashSet<OpenXmlPart> VisitedXmlParts { get; } = new();
        public long XmlBytesRead { get; set; }
    }

    public IrDrawingGraphHasher(
        OpenXmlPart drawingOwner,
        Lazy<IrHash> sourceDocumentFingerprint,
        IrDrawingGraphHashLimits? limits = null)
    {
        _drawingOwner = drawingOwner ?? throw new ArgumentNullException(nameof(drawingOwner));
        _sourceDocumentFingerprint = sourceDocumentFingerprint
            ?? throw new ArgumentNullException(nameof(sourceDocumentFingerprint));
        _limits = limits ?? IrDrawingGraphHashLimits.Default;
        if (_limits.MaxXmlGraphDepth < 0 || _limits.MaxXmlGraphParts < 0 ||
            _limits.MaxXmlPartBytes <= 0 || _limits.MaxXmlGraphBytes <= 0)
            throw new ArgumentOutOfRangeException(nameof(limits), "Drawing graph limits must be non-negative and byte limits positive.");
    }

    /// <summary>
    /// Hash one outer <c>w:drawing</c>. The shared canonicalizer retains all current drawing normalization
    /// (rsids, PowerTools data, nonvisual drawing ids, attribute ordering); only relationship values are
    /// replaced with graph-aware tokens.
    /// </summary>
    public IrHash Hash(XElement drawing)
    {
        ArgumentNullException.ThrowIfNull(drawing);
        var traversal = new Traversal();
        return IrHasher.CanonicalHashWithAttributeRewrite(
            drawing, attribute => RewriteAttribute(
                _drawingOwner, diagramDataRelationshipOwner: null, depth: 0, attribute, traversal));
    }

    /// <summary>
    /// True when this outer drawing directly starts a supported XML relationship graph. A graph must begin at
    /// the drawing owner, so this inexpensive check deliberately does not walk descendants of target parts.
    /// It lets textbox modeling retain an outer structural carrier only when the graph-aware identity adds
    /// information that the textbox's inner-block model cannot represent.
    /// </summary>
    public bool HasReachableXmlGraph(XElement drawing)
    {
        ArgumentNullException.ThrowIfNull(drawing);
        foreach (var attribute in drawing.DescendantsAndSelf().Attributes())
        {
            if (ImportedRelationshipAttributes.Contains(attribute.Name) &&
                TryGetInternalPart(_drawingOwner, attribute.Value, out var target) && IsXml(target))
                return true;
        }
        return false;
    }

    private XAttribute RewriteAttribute(
        OpenXmlPart owner, OpenXmlPart? diagramDataRelationshipOwner, int depth, XAttribute attribute,
        Traversal traversal,
        Action<bool>? reportCacheability = null)
    {
        if (ImportedRelationshipAttributes.Contains(attribute.Name))
        {
            var resolution = ResolveImportedRelationship(
                owner, attribute.Value, depth, traversal, graphScoped: !ReferenceEquals(owner, _drawingOwner));
            reportCacheability?.Invoke(resolution.Cacheable);
            return new XAttribute(attribute.Name, resolution.Value);
        }

        // SmartArt's data-model extension has an unqualified relId whose relationship is owned by the part that
        // linked the DiagramData part, not the DiagramData part that contains the XML. The renderer follows the
        // same special edge when importing a prebuilt diagram drawing. Do not treat a coincidental
        // dsp:dataModelExt in arbitrary XML as graph syntax.
        if (attribute.Name.NamespaceName.Length == 0 && attribute.Name.LocalName == "relId" &&
            attribute.Parent?.Name == DspDataModelExt && IsDiagramDataPart(owner))
        {
            var resolution = ResolveImportedRelationship(
                diagramDataRelationshipOwner ?? _drawingOwner,
                attribute.Value, depth, traversal, graphScoped: true);
            reportCacheability?.Invoke(resolution.Cacheable);
            return new XAttribute(attribute.Name, resolution.Value);
        }

        // Preserve the existing renumbering-neutral treatment of every other relationship-namespace
        // attribute, but deliberately do not add unsupported XML targets to this graph.
        if (attribute.Name.Namespace == R)
            return new XAttribute(attribute.Name, LegacyResolver(owner).ResolveToken(attribute.Value));

        return attribute;
    }

    private GraphIdentity ResolveImportedRelationship(
        OpenXmlPart owner, string relationshipId, int depth, Traversal traversal, bool graphScoped)
    {
        if (TryGetInternalPart(owner, relationshipId, out var target))
        {
            if (IsXml(target))
            {
                var identity = XmlPartIdentity(target, owner, depth + 1, traversal);
                return new GraphIdentity(RelationshipToken(target, identity.Value), identity.Cacheable);
            }

            var legacy = LegacyResolver(owner).ResolveToken(relationshipId);
            return graphScoped
                ? new GraphIdentity(BinaryRelationshipToken(target, legacy), Cacheable: true)
                : new GraphIdentity(legacy, Cacheable: true);
        }

        if (TryGetExternalRelationship(owner, relationshipId, out var relationshipType, out var targetUri))
        {
            var legacy = LegacyResolver(owner).ResolveToken(relationshipId);
            return graphScoped
                ? new GraphIdentity(ExternalRelationshipToken(relationshipType, targetUri), Cacheable: true)
                : new GraphIdentity(legacy, Cacheable: true);
        }

        return new GraphIdentity(LegacyResolver(owner).ResolveToken(relationshipId), Cacheable: true);
    }

    private GraphIdentity XmlPartIdentity(
        OpenXmlPart part, OpenXmlPart relationshipOwner, int depth, Traversal traversal)
    {
        if (depth > _limits.MaxXmlGraphDepth)
            return IncompleteIdentity(part, "depth");

        var cacheKey = new XmlCacheKey(
            part, IsDiagramDataPart(part) ? relationshipOwner : null, _limits.MaxXmlGraphDepth - depth);
        if (_xmlIdentityCache.TryGetValue(cacheKey, out var cached))
            return new GraphIdentity(cached, Cacheable: true);

        // Relationship graphs should be acyclic, but corrupt packages need deterministic total behavior. The
        // conservative cutoff binds the result to this source package rather than allowing an incomplete cycle
        // walk to make two different graphs compare Equal.
        if (!traversal.ActiveXmlParts.Add(part))
            return IncompleteIdentity(part, "cycle");

        string identity;
        bool cacheable = true;
        try
        {
            if (traversal.VisitedXmlParts.Add(part) &&
                traversal.VisitedXmlParts.Count > _limits.MaxXmlGraphParts)
            {
                identity = IncompleteIdentity(part, "part-limit").Value;
                cacheable = false;
            }
            else
            {
                var read = ReadXmlBytes(part, traversal);
                if (read.LimitReached)
                {
                    identity = IncompleteIdentity(part, read.LimitReason!).Value;
                    cacheable = false;
                }
                else if (read.Bytes!.Length == 0)
                {
                    identity = TypedLeaf("empty", part.ContentType);
                }
                else
                {
                    try
                    {
                        using var stream = new MemoryStream(read.Bytes, writable: false);
                        var root = XDocument.Load(stream).Root;
                        if (root is null)
                        {
                            identity = TypedLeaf("empty", part.ContentType);
                        }
                        else
                        {
                            bool childCacheable = true;
                            var canonical = IrHasher.CanonicalHashWithAttributeRewrite(
                                root,
                                attribute => RewriteAttribute(
                                    part, relationshipOwner, depth, attribute, traversal,
                                    canCache => childCacheable &= canCache));
                            identity = "sha:" + HashFramedXml(part.ContentType, canonical).ToHex();
                            cacheable = childCacheable;
                        }
                    }
                    catch (Exception e) when (e is XmlException or ArgumentException)
                    {
                        // The source bytes were readable but not XML. Keep an identity for those bytes rather
                        // than collapsing every malformed chart/diagram of one content type to Equal. The
                        // importer copies the raw part and deliberately skips recursive XML fixup in this case.
                        identity = "sha:" + HashFramedBytes(
                            "drawing-xml-malformed/v1", part.ContentType, read.Bytes).ToHex();
                    }
                }
            }
        }
        catch (Exception e) when (e is IOException or InvalidOperationException or NotSupportedException
            or ObjectDisposedException or UnauthorizedAccessException or KeyNotFoundException)
        {
            identity = IncompleteIdentityWithoutRawPart(part, "unreadable").Value;
            cacheable = false;
        }
        finally
        {
            traversal.ActiveXmlParts.Remove(part);
        }

        if (cacheable)
            _xmlIdentityCache[cacheKey] = identity;
        return new GraphIdentity(identity, cacheable);
    }

    private XmlReadResult ReadXmlBytes(OpenXmlPart part, Traversal traversal)
    {
        if (traversal.XmlBytesRead >= _limits.MaxXmlGraphBytes)
            return new XmlReadResult(null, "graph-byte-limit");

        using var source = part.GetStream(FileMode.Open, FileAccess.Read);
        using var buffer = new MemoryStream();
        var chunk = new byte[81920];
        while (true)
        {
            int read = source.Read(chunk, 0, chunk.Length);
            if (read == 0)
                break;

            if (buffer.Length + read > _limits.MaxXmlPartBytes)
            {
                traversal.XmlBytesRead = _limits.MaxXmlGraphBytes;
                return new XmlReadResult(null, "part-byte-limit");
            }
            if (traversal.XmlBytesRead + buffer.Length + read > _limits.MaxXmlGraphBytes)
            {
                traversal.XmlBytesRead = _limits.MaxXmlGraphBytes;
                return new XmlReadResult(null, "graph-byte-limit");
            }
            buffer.Write(chunk, 0, read);
        }

        traversal.XmlBytesRead += buffer.Length;
        return new XmlReadResult(buffer.ToArray(), LimitReason: null);
    }

    /// <summary>
    /// A graph that hit a traversal limit is deliberately incomparable across distinct source packages. Its
    /// immediate part still receives a full streaming raw-byte SHA-256 (no large allocation), while the source
    /// package fingerprint closes the remaining hole: an unchanged oversized chart XML can still point at an
    /// embedded workbook or nested XML part whose content changed beyond the cutoff. This may over-report a
    /// drawing change when unrelated package bytes differ, but it never lets a partially inspected graph align
    /// Equal and leak the right graph through Reject.
    /// </summary>
    private GraphIdentity IncompleteIdentity(OpenXmlPart part, string reason)
    {
        try
        {
            return BuildIncompleteIdentity(part, reason, GetRawPartIdentity(part));
        }
        catch (Exception e) when (e is IOException or InvalidOperationException or NotSupportedException
            or ObjectDisposedException or UnauthorizedAccessException or KeyNotFoundException)
        {
            return IncompleteIdentityWithoutRawPart(part, reason);
        }
    }

    private GraphIdentity IncompleteIdentityWithoutRawPart(OpenXmlPart part, string reason) =>
        BuildIncompleteIdentity(part, reason, rawPart: null);

    private GraphIdentity BuildIncompleteIdentity(
        OpenXmlPart part, string reason, RawPartIdentity? rawPart)
    {
        using var stream = new MemoryStream();
        WriteFrame(stream, "drawing-xml-incomplete/v1");
        WriteFrame(stream, reason);
        WriteFrame(stream, part.ContentType);
        if (rawPart is { } raw)
        {
            stream.WriteByte(1);
            WriteInt64(stream, raw.Length);
            stream.Write(raw.Digest, 0, raw.Digest.Length);
        }
        else
        {
            stream.WriteByte(0);
        }

        // The fallback is source-package scoped rather than URI/rId scoped. rId and package part names churn
        // freely across equivalent documents, but a cutoff cannot establish equivalence, so conservative false
        // positives are preferable to carrying an uninspected child relationship from the right source into a
        // rejected redline.
        WriteFrame(stream, "drawing-xml-source-package/v1");
        Span<byte> sourceDigest = stackalloc byte[32];
        _sourceDocumentFingerprint.Value.CopyTo(sourceDigest);
        stream.Write(sourceDigest);
        return new GraphIdentity(
            "sha:" + IrHash.Compute(stream.GetBuffer().AsSpan(0, (int)stream.Length)).ToHex(),
            Cacheable: false);
    }

    private RawPartIdentity GetRawPartIdentity(OpenXmlPart part)
    {
        if (_rawPartIdentityCache.TryGetValue(part, out var cached))
            return cached;

        using var source = part.GetStream(FileMode.Open, FileAccess.Read);
        using var hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        var chunk = new byte[81920];
        long length = 0;
        while (true)
        {
            int read = source.Read(chunk, 0, chunk.Length);
            if (read == 0)
                break;
            if (length > long.MaxValue - read)
                throw new IOException("Drawing XML part length exceeded the supported range.");

            hash.AppendData(chunk, 0, read);
            length += read;
        }

        var identity = new RawPartIdentity(length, hash.GetHashAndReset());
        _rawPartIdentityCache.Add(part, identity);
        return identity;
    }

    private static IrHash HashFramedXml(string contentType, IrHash canonicalXml)
    {
        using var stream = new MemoryStream();
        WriteFrame(stream, "drawing-xml-graph/v1");
        WriteFrame(stream, contentType);
        Span<byte> digest = stackalloc byte[32];
        canonicalXml.CopyTo(digest);
        stream.Write(digest);
        return IrHash.Compute(stream.GetBuffer().AsSpan(0, (int)stream.Length));
    }

    private static IrHash HashFramedBytes(string kind, string contentType, byte[] bytes)
    {
        using var stream = new MemoryStream();
        WriteFrame(stream, kind);
        WriteFrame(stream, contentType);
        stream.Write(bytes, 0, bytes.Length);
        return IrHash.Compute(stream.GetBuffer().AsSpan(0, (int)stream.Length));
    }

    private static string RelationshipToken(OpenXmlPart target, string identity) =>
        "drawing-xml-rel/v1|" + Frame(target.RelationshipType) + "|" + Frame(target.ContentType) + "|" + identity;

    private static string BinaryRelationshipToken(OpenXmlPart target, string legacyIdentity) =>
        "drawing-binary-rel/v1|" + Frame(target.RelationshipType) + "|" + Frame(target.ContentType) + "|" +
        Frame(legacyIdentity);

    private static string ExternalRelationshipToken(string relationshipType, Uri targetUri) =>
        "drawing-external-rel/v1|" + Frame(relationshipType) + "|" + Frame(targetUri.ToString());

    private static string TypedLeaf(string kind, string contentType) =>
        "drawing-xml-" + kind + "/v1|" + Frame(contentType);

    private static string Frame(string? value)
    {
        value ??= string.Empty;
        return value.Length.ToString(CultureInfo.InvariantCulture) + ":" + value;
    }

    private static void WriteFrame(Stream stream, string value)
    {
        var bytes = Encoding.UTF8.GetBytes(value);
        Span<byte> length = stackalloc byte[4];
        BinaryPrimitives.WriteInt32BigEndian(length, bytes.Length);
        stream.Write(length);
        stream.Write(bytes);
    }

    private static void WriteInt64(Stream stream, long value)
    {
        Span<byte> bytes = stackalloc byte[8];
        BinaryPrimitives.WriteInt64BigEndian(bytes, value);
        stream.Write(bytes);
    }

    private IrRelResolver LegacyResolver(OpenXmlPart owner)
    {
        if (_legacyResolvers.TryGetValue(owner, out var resolver))
            return resolver;
        resolver = new IrRelResolver(owner);
        _legacyResolvers.Add(owner, resolver);
        return resolver;
    }

    private static bool TryGetInternalPart(OpenXmlPart owner, string relationshipId, out OpenXmlPart target)
    {
        try
        {
            foreach (var pair in owner.Parts)
                if (pair.RelationshipId == relationshipId)
                {
                    target = pair.OpenXmlPart;
                    return true;
                }
        }
        catch (Exception e) when (e is InvalidOperationException or NotSupportedException
            or ObjectDisposedException or ArgumentException or KeyNotFoundException)
        {
            // Fall through to the legacy totality token below.
        }

        target = null!;
        return false;
    }

    private static bool TryGetExternalRelationship(
        OpenXmlPart owner, string relationshipId, out string relationshipType, out Uri targetUri)
    {
        try
        {
            var external = owner.ExternalRelationships.FirstOrDefault(r => r.Id == relationshipId);
            if (external is not null && external.Uri is not null)
            {
                relationshipType = external.RelationshipType;
                targetUri = external.Uri;
                return true;
            }

            var hyperlink = owner.HyperlinkRelationships.FirstOrDefault(r => r.Id == relationshipId);
            if (hyperlink is not null && hyperlink.Uri is not null)
            {
                relationshipType = hyperlink.RelationshipType;
                targetUri = hyperlink.Uri;
                return true;
            }
        }
        catch (Exception e) when (e is InvalidOperationException or NotSupportedException
            or ObjectDisposedException or ArgumentException or KeyNotFoundException)
        {
            // Fall through to the legacy totality token below.
        }

        relationshipType = string.Empty;
        targetUri = null!;
        return false;
    }

    private static bool IsXml(OpenXmlPart part) =>
        part.ContentType.EndsWith("xml", StringComparison.OrdinalIgnoreCase);

    private static bool IsDiagramDataPart(OpenXmlPart part) =>
        part.RelationshipType.EndsWith("/diagramData", StringComparison.Ordinal);
}
