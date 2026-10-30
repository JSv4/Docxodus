using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.IO.Packaging;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Xml;
using Docxodus.Verification;
using DocumentFormat.OpenXml.Packaging;
using PackageRelationship = System.IO.Packaging.PackageRelationship;

namespace Docxodus.Internal;

/// <summary>
/// The XML depth door (issue #1389). The shared read step (<see cref="StrictOoxmlNormalizer"/>)
/// measures every part of a package that a later step can build a tree from — every part the package
/// relationships reach, as <see cref="System.IO.Packaging"/> assembles it (interleaved pieces
/// included), which is what the SDK loads — streaming each once through
/// <see cref="DepthLimitedXmlReader"/> before any tree is built. A part nested past its limit refuses
/// the package, so no recursive walk (part snapshots, the SDK's typed DOM, the HTML render,
/// comparison, revisions) meets an over-deep tree. A WordprocessingML story or definition part —
/// reached by a story relationship or declared a story content type — is held to
/// <see cref="XmlSemanticNormalizer.MaxStoryDepth"/>; every other part to
/// <see cref="XmlSemanticNormalizer.MaxElementDepth"/>. Parts nothing reaches are never parsed by a
/// later step and are not read here either. A package a caller opened never passes the read step, so
/// each public entry point that takes an open document measures it through <see cref="Enter"/>; a
/// session's edits are held to the same limits by <see cref="PartDepthWatch"/>.
/// </summary>
internal static class PartDepthGuard
{
    // Byte arrays already measured, with a hash of their content: the read step runs several times
    // per pipeline over the same bytes, and a hit counts only while the content is what was measured.
    private static readonly ConditionalWeakTable<byte[], StrongBox<ulong>> Scanned = new();

    // The package manifest's raw-boundary limits bound the work: what a scan may read, per part and in
    // all, and how far a part may inflate (issue #1389).
    private static readonly PackageManifestOptions Limits = new();

    // The scan expands a DTD the way the most lenient later tree builder (XDocument.Load, LINQ to
    // XML's settings) does, so it measures the tree that builder would build. Word never writes a DTD;
    // one part may carry one, and a second refuses the package rather than multiply that work.
    private static readonly XmlReaderSettings Lenient = new()
    {
        DtdProcessing = DtdProcessing.Parse,
        MaxCharactersFromEntities = 10_000_000,
        XmlResolver = null,
        // Whitespace adds no depth; skipping it keeps a whitespace-only run from being buffered.
        IgnoreWhitespace = true,
    };

    private static readonly HashSet<string> StoryRelationshipTypes = new(StringComparer.Ordinal)
    {
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
        "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
        "http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument",
    };

    private static readonly HashSet<string> StoryContentTypes = new(StringComparer.OrdinalIgnoreCase)
    {
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml",
        "application/vnd.ms-word.document.macroEnabled.main+xml",
        "application/vnd.ms-word.template.macroEnabledTemplate.main+xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
        "application/vnd.ms-word.stylesWithEffects+xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
    };

    /// <summary>Throws <see cref="XmlDepthLimitException"/> when a part a later step can parse nests
    /// deeper than its limit, or when measuring the package would pass the manifest's raw-boundary
    /// limits. A package the packaging layer cannot open is left for the open that follows to report:
    /// nothing builds a tree from it.</summary>
    internal static void EnsurePackage(byte[] packageBytes)
    {
        var hash = ContentHash(packageBytes);
        if (Scanned.TryGetValue(packageBytes, out var seen) && seen.Value == hash) return;
        Package package;
        try
        {
            package = Package.Open(new MemoryStream(packageBytes, writable: false), FileMode.Open, FileAccess.Read);
        }
        catch (Exception e) when (e is not XmlDepthLimitException and not OutOfMemoryException) { return; }
        using (package)
        {
            var sizes = EntrySizes(packageBytes);
            long read = 0;
            bool dtdSeen = false;
            foreach (var (part, story) in ReachableParts(package))
            {
                var (compressed, uncompressed) = sizes(part.Uri.OriginalString);
                if (uncompressed > Limits.MaxEntryUncompressedBytes)
                    Refuse("entry_size_limit_exceeded", part.Uri.OriginalString);
                if (compressed > 0 && (double)uncompressed / compressed > Limits.MaxCompressionRatio)
                    Refuse("compression_ratio_limit_exceeded", part.Uri.OriginalString);
                if ((read += uncompressed) > Limits.MaxTotalUncompressedBytes)
                    Refuse("total_expansion_limit_exceeded", part.Uri.OriginalString);
                Scan(part.Uri.OriginalString, () => part.GetStream(FileMode.Open, FileAccess.Read),
                    story ? XmlSemanticNormalizer.MaxStoryDepth : XmlSemanticNormalizer.MaxElementDepth, ref dtdSeen);
            }
        }
        Scanned.AddOrUpdate(packageBytes, new StrongBox<ulong>(hash));
    }

    // Open packages the library opened from bytes the door already measured, and ones an entry point
    // is measuring for the call in progress, so the entry points it calls in turn skip them.
    private static readonly ConditionalWeakTable<OpenXmlPackage, Mark> Measured = new();

    /// <summary>How many open packages this thread has read part streams of to measure them; tests read
    /// it to prove an entry point reads a caller's package once and the library's own never.</summary>
    [ThreadStatic]
    internal static int OpenPackageMeasurements;

    /// <summary>Record that <paramref name="package"/> was opened from bytes <see cref="EnsurePackage"/>
    /// measured, so only trees loaded from it since are measured; <paramref name="watched"/> when every
    /// change to its trees is measured as it happens (<see cref="PartDepthWatch"/>), so nothing is.</summary>
    internal static void MarkMeasured(OpenXmlPackage package, bool watched = false)
    {
        var mark = Measured.GetOrCreateValue(package);
        mark.StreamsMeasured = true;
        mark.Watched |= watched;
    }

    /// <summary>The door for a package a caller opened (issue #1389): measures every part its
    /// relationships reach — the part's loaded LINQ tree if it has one, otherwise its stream, plus any
    /// loaded SDK root, which a later step may read or serialize. The streams of a package the library
    /// opened from measured bytes are not read again, and a watched package or one an entry point
    /// already measures for the call in progress is skipped. Each public entry point that takes an
    /// open document holds the returned scope for its call. An open package exposes no entry sizes,
    /// so the manifest's size and ratio limits do not apply here; the scan reads what the call would read.</summary>
    /// <param name="preferLoadedSdkRoots">The caller reads a loaded SDK root instead of its stream
    /// whenever one exists (the OpenContracts exporter). Such a stream need not be scanned; opening
    /// an empty stream through the SDK would also discard its newly assigned, unsaved root.</param>
    internal static Scope Enter(OpenXmlPackage package, bool preferLoadedSdkRoots = false)
    {
        var mark = Measured.GetOrCreateValue(package);
        if (!mark.Watched && mark.Scopes == 0)
            EnsureOpenPackage(package, readStreams: !mark.StreamsMeasured, preferLoadedSdkRoots);
        mark.Scopes++;
        return new Scope(mark);
    }

    /// <summary><see cref="Enter"/> for an entry point that walks one part recursively and reads the rest
    /// flat: measures that part alone, so a light query does not pay for the whole package. A part not
    /// yet loaded is loaded here through the depth-limited reader, which is the tree the call reads, so
    /// the measurement costs no second parse.</summary>
    internal static void EnsurePart(OpenXmlPart part)
    {
        var mark = Measured.GetOrCreateValue(part.OpenXmlPackage);
        if (mark.Watched || mark.Scopes > 0) return;
        var limit = LimitFor(part);
        if (part.Annotation<System.Xml.Linq.XDocument>() is not null || mark.StreamsMeasured)
        {
            bool dtdSeen = false;
            Measure(part, readStream: false, ref dtdSeen);
            return;
        }
        OpenPackageMeasurements++;
        try { part.GetXDocument(limit); }
        catch (XmlDepthLimitException e) when (e.PartUri is null) { throw DepthRefusal(part.Uri.OriginalString, limit); }
    }

    private static void EnsureOpenPackage(OpenXmlPackage package, bool readStreams, bool preferLoadedSdkRoots)
    {
        if (readStreams) OpenPackageMeasurements++;
        bool dtdSeen = false;
        var seen = new HashSet<OpenXmlPart>(ReferenceEqualityComparer.Instance);
        var pending = new Stack<OpenXmlPartContainer>();
        pending.Push(package);
        while (pending.Count > 0)
        {
            foreach (var pair in pending.Pop().Parts)
            {
                if (!seen.Add(pair.OpenXmlPart)) continue;
                pending.Push(pair.OpenXmlPart);
                Measure(pair.OpenXmlPart, readStreams, ref dtdSeen, preferLoadedSdkRoots);
            }
        }
    }

    // LINQ and SDK trees can coexist and disagree: measure both loaded representations before a
    // caller reads or serializes either. Checking IsRootElementLoaded avoids loading an unchecked
    // stream through the SDK's recursive parser when only its LINQ tree is in use.
    private static void Measure(OpenXmlPart part, bool readStream, ref bool dtdSeen, bool preferLoadedSdkRoots = false)
    {
        var limit = LimitFor(part);
        var partUri = part.Uri.OriginalString;
        var sdkRoot = part.IsRootElementLoaded ? part.RootElement : null;
        if (sdkRoot is not null)
        {
            // ChildElements can force a lazy SDK node to parse arbitrarily deep raw XML. WriteTo
            // instead exposes raw nodes as WriteRaw and materialized nodes one element at a time,
            // so both forms stop at the limit before parsing or serializing a deeper subtree.
            using var writer = new DepthMeasuringXmlWriter(limit);
            try { writer.Measure(sdkRoot); }
            catch (XmlDepthLimitException) { throw DepthRefusal(partUri, limit); }
        }
        if (part.Annotation<System.Xml.Linq.XDocument>() is { } tree)
        {
            try { XmlSemanticNormalizer.EnsureBoundedDepth(tree, 0, limit); }
            catch (XmlDepthLimitException) { throw DepthRefusal(partUri, limit); }
        }
        else if (readStream && !(preferLoadedSdkRoots && sdkRoot is not null))
            Scan(partUri, () => part.GetStream(FileMode.Open, FileAccess.Read), limit, ref dtdSeen);
    }

    internal sealed class Mark
    {
        internal bool StreamsMeasured;
        internal bool Watched;
        internal int Scopes;
    }

    /// <summary>Held by an entry point for its call; see <see cref="Enter"/>.</summary>
    internal readonly struct Scope(Mark mark) : IDisposable
    {
        public void Dispose() => mark.Scopes--;
    }

    private static void Scan(string partUri, Func<Stream> open, int limit, ref bool dtdSeen)
    {
        try
        {
            using var stream = open();
            using var reader = new DepthLimitedXmlReader(XmlReader.Create(stream, Lenient), limit);
            while (reader.Read())
            {
                if (reader.NodeType != XmlNodeType.DocumentType) continue;
                if (dtdSeen) Refuse("xml_size_limit_exceeded", partUri);
                dtdSeen = true;
            }
        }
        catch (XmlDepthLimitException e) when (e.PartUri is null)
        {
            throw DepthRefusal(partUri, limit);
        }
        // Not XML, XML no reader accepts, or entities past LINQ to XML's own cap: no later step builds
        // a tree from it either.
        catch (XmlException e) when (e is not XmlDepthLimitException) { }
        catch (Exception e) when (e is IOException or InvalidDataException or NotSupportedException) { }
    }

    private static XmlDepthLimitException DepthRefusal(string partUri, int limit) => new(
        $"xml_depth_limit_exceeded: {partUri} nests deeper than the {limit} level safety limit.")
    { PartUri = partUri };

    /// <summary>Every part the package relationships reach, transitively, each once, with whether a
    /// story relationship reaches it or it is declared a story content type.</summary>
    private static IEnumerable<(PackagePart Part, bool Story)> ReachableParts(Package package)
    {
        // Keyed by the part's own name, which the packaging layer matches without regard to case.
        var story = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        var order = new List<PackagePart>();
        var pending = new Queue<(Uri Source, Func<PackageRelationshipCollection> Relationships)>();
        pending.Enqueue((new Uri("/", UriKind.Relative), package.GetRelationships));
        while (pending.Count > 0)
        {
            var (source, relationships) = pending.Dequeue();
            List<PackageRelationship> list;
            // Relationships the packaging layer cannot read (an unparseable target, say) are read
            // again after open-time repair rewrites them, when the read step measures the repair.
            try { list = new List<PackageRelationship>(relationships()); }
            catch (Exception e) when (e is not XmlDepthLimitException and not OutOfMemoryException) { continue; }
            foreach (var relationship in list)
            {
                if (relationship.TargetMode != TargetMode.Internal) continue;
                PackagePart part;
                try
                {
                    var target = PackUriHelper.ResolvePartUri(source, relationship.TargetUri);
                    if (!package.PartExists(target)) continue;
                    part = package.GetPart(target);
                }
                catch (Exception e) when (e is not XmlDepthLimitException and not OutOfMemoryException) { continue; }
                var name = part.Uri.OriginalString;
                if (story.TryGetValue(name, out var known))
                {
                    story[name] = known || StoryRelationshipTypes.Contains(relationship.RelationshipType);
                    continue;
                }
                story[name] = IsStory(relationship.RelationshipType, part.ContentType);
                order.Add(part);
                pending.Enqueue((part.Uri, part.GetRelationships));
            }
        }
        foreach (var part in order) yield return (part, story[part.Uri.OriginalString]);
    }

    /// <summary>The depth limit of a part a story relationship reaches or a story content type declares.</summary>
    internal static int LimitFor(OpenXmlPart part) =>
        IsStory(part.RelationshipType, part.ContentType) ? XmlSemanticNormalizer.MaxStoryDepth : XmlSemanticNormalizer.MaxElementDepth;

    private static bool IsStory(string relationshipType, string contentType) =>
        StoryRelationshipTypes.Contains(relationshipType) || StoryContentTypes.Contains(contentType.Split(';')[0].Trim());

    /// <summary>Compressed and uncompressed size of a part, summing its interleaved pieces.</summary>
    private static Func<string, (long Compressed, long Uncompressed)> EntrySizes(byte[] packageBytes)
    {
        var sizes = new Dictionary<string, (long, long)>(StringComparer.OrdinalIgnoreCase);
        using var zip = new ZipArchive(new MemoryStream(packageBytes, writable: false), ZipArchiveMode.Read);
        foreach (var entry in zip.Entries)
        {
            var name = "/" + PackageOpenRepair.WithoutPieceSuffix(entry.FullName);
            sizes.TryGetValue(name, out var size);
            sizes[name] = (size.Item1 + entry.CompressedLength, size.Item2 + entry.Length);
        }
        return name => sizes.TryGetValue(Uri.UnescapeDataString(name), out var size) ? size : (0, 0);
    }

    private static void Refuse(string code, string partUri) => throw new XmlDepthLimitException(
        $"{code}: measuring {partUri} for the XML depth limit would pass the package limits.")
    { PartUri = partUri, IsBudgetRefusal = true };

    private static ulong ContentHash(byte[] bytes)
    {
        ulong hash = 14695981039346656037UL ^ (ulong)bytes.Length;
        var words = MemoryMarshal.Cast<byte, ulong>(bytes);
        foreach (var word in words) hash = (hash ^ word) * 1099511628211UL;
        for (int i = words.Length * sizeof(ulong); i < bytes.Length; i++) hash = (hash ^ bytes[i]) * 1099511628211UL;
        return hash;
    }
}

/// <summary>
/// Forwards an <see cref="XmlReader"/> and throws <see cref="XmlDepthLimitException"/> as soon as it
/// reads an element nested deeper than <paramref name="limit"/>, having read no more than that many
/// levels.
/// </summary>
internal sealed class DepthLimitedXmlReader(XmlReader inner, int limit) : XmlReader
{
    /// <summary><see cref="System.Xml.Linq.XDocument.Parse(string, System.Xml.Linq.LoadOptions)"/>, refusing past
    /// <see cref="XmlSemanticNormalizer.MaxElementDepth"/>: for steps that parse package entries themselves,
    /// whatever the read step measured, and then walk the tree recursively (issue #1389).</summary>
    internal static System.Xml.Linq.XDocument Parse(string xml, System.Xml.Linq.LoadOptions options)
    {
        using var reader = Create(xml, options);
        return System.Xml.Linq.XDocument.Load(reader, options);
    }

    /// <summary><see cref="System.Xml.Linq.XElement.Parse(string)"/>, refusing past
    /// <see cref="XmlSemanticNormalizer.MaxElementDepth"/> (issue #1389).</summary>
    internal static System.Xml.Linq.XElement ParseElement(string xml)
    {
        using var reader = Create(xml, System.Xml.Linq.LoadOptions.None);
        return System.Xml.Linq.XElement.Load(reader);
    }

    // The reader settings LINQ to XML's own Parse uses for these options.
    private static DepthLimitedXmlReader Create(string xml, System.Xml.Linq.LoadOptions options) => new(
        XmlReader.Create(new StringReader(xml), new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Parse,
            MaxCharactersFromEntities = 10_000_000,
            IgnoreWhitespace = (options & System.Xml.Linq.LoadOptions.PreserveWhitespace) == 0,
        }),
        XmlSemanticNormalizer.MaxElementDepth);

    public override bool Read()
    {
        if (!inner.Read()) return false;
        // Depth is zero-based: the root element is at 0 and nests one level deep.
        if (inner.NodeType == XmlNodeType.Element && inner.Depth >= limit) XmlSemanticNormalizer.ThrowDepthLimit(limit);
        return true;
    }

    public override int AttributeCount => inner.AttributeCount;
    public override string BaseURI => inner.BaseURI;
    public override int Depth => inner.Depth;
    public override bool EOF => inner.EOF;
    public override bool HasValue => inner.HasValue;
    public override bool IsDefault => inner.IsDefault;
    public override bool IsEmptyElement => inner.IsEmptyElement;
    public override string LocalName => inner.LocalName;
    public override string Name => inner.Name;
    public override string NamespaceURI => inner.NamespaceURI;
    public override XmlNameTable NameTable => inner.NameTable;
    public override XmlNodeType NodeType => inner.NodeType;
    public override string Prefix => inner.Prefix;
    public override ReadState ReadState => inner.ReadState;
    public override XmlReaderSettings? Settings => inner.Settings;
    public override string Value => inner.Value;
    public override string XmlLang => inner.XmlLang;
    public override XmlSpace XmlSpace => inner.XmlSpace;
    public override string GetAttribute(int i) => inner.GetAttribute(i);
    public override string? GetAttribute(string name) => inner.GetAttribute(name);
    public override string? GetAttribute(string name, string? namespaceURI) => inner.GetAttribute(name, namespaceURI);
    public override string? LookupNamespace(string prefix) => inner.LookupNamespace(prefix);
    public override bool MoveToAttribute(string name) => inner.MoveToAttribute(name);
    public override bool MoveToAttribute(string name, string? ns) => inner.MoveToAttribute(name, ns);
    public override void MoveToAttribute(int i) => inner.MoveToAttribute(i);
    public override bool MoveToElement() => inner.MoveToElement();
    public override bool MoveToFirstAttribute() => inner.MoveToFirstAttribute();
    public override bool MoveToNextAttribute() => inner.MoveToNextAttribute();
    public override bool ReadAttributeValue() => inner.ReadAttributeValue();
    public override void ResolveEntity() => inner.ResolveEntity();
    public override bool CanResolveEntity => inner.CanResolveEntity;

    protected override void Dispose(bool disposing)
    {
        if (disposing) inner.Dispose();
        base.Dispose(disposing);
    }
}
