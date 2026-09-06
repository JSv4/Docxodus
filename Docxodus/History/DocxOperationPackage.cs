#nullable enable

using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Docxodus.Verification;

namespace Docxodus.History;

/// <summary>Conservative package/text reconciliation, not a live-session mutation recorder.</summary>
internal static class DocxOperationPackage
{
    private sealed record Entry(string Name, byte[] Bytes);
    private static readonly XNamespace Word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    internal static byte[] ApplyText(byte[] package, DocxTextSplice splice, PackageManifestOptions? options)
    {
        DocxOperationStore.Text(splice);
        var entries = Read(package, options);
        if (!entries.TryGetValue(splice.PartUri, out var entry)) throw Invalid("Text part is absent at the supplied base.");
        using var input = new MemoryStream(entry.Bytes, writable: false);
        using var reader = XmlReader.Create(input, new XmlReaderSettings
        { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = entry.Bytes.Length + 1L });
        var document = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        var node = document.Descendants(Word + "t").Skip(splice.TextNode).FirstOrDefault();
        if (node is null || node.Nodes().Any(n => n is not XText)) throw Invalid("Text node is absent or has unsupported children.");
        var text = node.Value;
        var end = (long)splice.Offset + splice.DeleteCount;
        if (end > text.Length || !Boundary(text, splice.Offset) || !Boundary(text, (int)end))
            throw Invalid("Splice must address valid UTF-16 boundaries in the base text.");
        var replacement = text[..splice.Offset] + splice.Insert + text[(int)end..];
        if (text == replacement) return package.ToArray();
        node.Value = replacement;
        node.SetAttributeValue(XNamespace.Xml + "space", "preserve");
        using var output = new MemoryStream();
        using (var writer = XmlWriter.Create(output, new XmlWriterSettings
        { Encoding = new UTF8Encoding(false, true), Indent = false, CloseOutput = false, NewLineHandling = NewLineHandling.Entitize }))
            document.Save(writer);
        entries[splice.PartUri] = entry with { Bytes = output.ToArray() };
        return Write(entries);
    }

    // Preserve accepted order for same-gap inserts. Never extend a stale deletion into text
    // inserted concurrently inside its observed range; contested overlaps remain explicit.
    internal static DocxTextSplice? Map(DocxTextSplice incoming, DocxTextSplice accepted)
    {
        if (incoming.PartUri != accepted.PartUri || incoming.TextNode != accepted.TextNode) return incoming;
        var p = incoming.Offset; var q = (long)p + incoming.DeleteCount;
        var a = accepted.Offset; var b = (long)a + accepted.DeleteCount;
        if (accepted.DeleteCount == 0)
        {
            if (q <= a && p < a) return incoming;
            if (p >= a) return incoming with { Offset = checked(p + accepted.Insert.Length) };
            return null;
        }
        if (q <= a) return incoming;
        if (p >= b) return incoming with { Offset = checked(p + accepted.Insert.Length - accepted.DeleteCount) };
        return null;
    }

    internal static bool IsTopology(string uri) => uri == "/[Content_Types].xml" || uri.EndsWith(".rels", StringComparison.Ordinal);

    internal static string? CheckReads(byte[] baseline, byte[] current, IReadOnlyList<string> readParts,
        PackageManifestOptions? options)
    {
        var left = Manifest(baseline, options).Entries.ToDictionary(e => e.Uri, e => e.RawBytesDigest, StringComparer.Ordinal);
        var right = Manifest(current, options).Entries.ToDictionary(e => e.Uri, e => e.RawBytesDigest, StringComparer.Ordinal);
        foreach (var uri in readParts.Concat(left.Keys.Concat(right.Keys).Where(IsTopology)).Distinct(StringComparer.Ordinal))
        {
            left.TryGetValue(uri, out var before); right.TryGetValue(uri, out var after);
            if (before != after) return "ReadDependencyChanged";
        }
        return null;
    }

    internal static (byte[]? Bytes, string? Conflict) Merge(byte[] current, PackageChangeSet proposed,
        PackageManifestOptions? options)
    {
        var identities = Manifest(current, options).Entries.ToDictionary(e => e.Uri, e => e.RawBytesDigest, StringComparer.Ordinal);
        foreach (var change in proposed.Changes)
        {
            identities.TryGetValue(change.Uri, out var actual);
            if (actual != change.BeforeDigest && actual != change.AfterDigest) return (null, "ChangedPart");
        }
        var entries = Read(current, options);
        var changed = false;
        foreach (var change in proposed.Changes)
        {
            identities.TryGetValue(change.Uri, out var actual);
            if (actual == change.AfterDigest) continue;
            changed = true;
            if (change.AfterDigest is null) entries.Remove(change.Uri);
            else entries[change.Uri] = new Entry(change.AfterEntryName!, proposed.GetPayload(change.AfterDigest));
        }
        return (changed ? Write(entries) : current.ToArray(), null);
    }

    private static Dictionary<string, Entry> Read(byte[] package, PackageManifestOptions? options)
    {
        _ = Manifest(package, options); // Bound expansion and validate OPC names before allocating entries.
        using var zip = new ZipArchive(new MemoryStream(package, writable: false), ZipArchiveMode.Read);
        var result = new Dictionary<string, Entry>(StringComparer.Ordinal);
        foreach (var entry in zip.Entries)
        {
            if (!PackageManifestGenerator.TryCanonicalizeEntryName(entry.FullName, out var uri)) throw Invalid("Invalid OPC part URI.");
            using var stream = entry.Open();
            var bytes = new byte[checked((int)entry.Length)]; stream.ReadExactly(bytes);
            result.Add(uri, new Entry(entry.FullName, bytes));
        }
        return result;
    }

    private static byte[] Write(Dictionary<string, Entry> entries)
    {
        using var output = new MemoryStream();
        using (var zip = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
            foreach (var pair in entries.OrderBy(e => e.Key, StringComparer.Ordinal))
            {
                var entry = zip.CreateEntry(pair.Value.Name, CompressionLevel.NoCompression);
                entry.LastWriteTime = new DateTimeOffset(1980, 1, 1, 0, 0, 0, TimeSpan.Zero);
                using var stream = entry.Open(); stream.Write(pair.Value.Bytes);
            }
        return output.ToArray();
    }

    private static PackageManifest Manifest(byte[] package, PackageManifestOptions? options)
    {
        var manifest = PackageManifestGenerator.Generate(package, options);
        if (!manifest.IsValid || manifest.PackageKind != "opc" || manifest.OrderedOpcContentDigest is null)
            throw Invalid("A valid bounded OPC package is required.");
        return manifest;
    }
    private static bool Boundary(string text, int at) => at >= 0 && at <= text.Length
        && (at == 0 || at == text.Length || !char.IsHighSurrogate(text[at - 1]) || !char.IsLowSurrogate(text[at]));
    private static PackageChangeException Invalid(string message) => new(PackageChangeError.InvalidPackage, message);
}
