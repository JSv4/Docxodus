// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace Docxodus.Verification;

/// <summary>
/// Narrow package fallback for semantic facts not represented by the IR. It is intentionally hidden
/// behind <see cref="ISemanticPackageChangeDetector"/> so the package manifest/delta from #456 can
/// replace entry reading, limits, normalized hashes, and relationship enumeration during rebase.
/// </summary>
internal sealed class OpcSemanticPackageChangeDetector : ISemanticPackageChangeDetector
{
    private static readonly HashSet<string> RevisionNames = new(StringComparer.Ordinal)
    {
        "ins", "del", "moveFrom", "moveTo", "moveFromRangeStart", "moveFromRangeEnd",
        "moveToRangeStart", "moveToRangeEnd", "rPrChange", "pPrChange", "tblPrChange",
        "tblGridChange", "trPrChange", "tblPrExChange", "tcPrChange", "sectPrChange",
        "numberingChange", "cellIns", "cellDel", "cellMerge", "customXmlInsRangeStart",
        "customXmlInsRangeEnd", "customXmlDelRangeStart", "customXmlDelRangeEnd",
        "customXmlMoveFromRangeStart", "customXmlMoveFromRangeEnd",
        "customXmlMoveToRangeStart", "customXmlMoveToRangeEnd",
    };

    private const string WordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string StrictWordNamespace = "http://purl.oclc.org/ooxml/wordprocessingml/main";
    private const string RelationshipNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string StrictRelationshipNamespace = "http://purl.oclc.org/ooxml/package/relationships";
    private const string OfficeRelationshipNamespace =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string StrictOfficeRelationshipNamespace =
        "http://purl.oclc.org/ooxml/officeDocument/relationships";
    private const string AnnotationNamespace = "http://docxodus.dev/annotations/v1";

    public IReadOnlyList<SemanticChangeDraft> Compare(
        byte[] leftBytes,
        byte[] rightBytes,
        SemanticDiffOptions options)
    {
        var left = Read(leftBytes, options);
        var right = Read(rightBytes, options);
        var result = new List<SemanticChangeDraft>();
        CompareEntities(left.Relationships, right.Relationships, result);
        CompareEntities(left.RelationshipBindings, right.RelationshipBindings, result);
        CompareEntities(left.Media, right.Media, result);
        CompareEntities(left.Bookmarks, right.Bookmarks, result);
        CompareEntities(left.Revisions, right.Revisions, result);
        CompareEntities(left.Annotations, right.Annotations, result);
        CompareEntities(left.RegistryParts, right.RegistryParts, result);
        CompareEntities(left.OpaqueParts, right.OpaqueParts, result);
        return result;
    }

    private static PackageSnapshot Read(byte[] bytes, SemanticDiffOptions options)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
        if (archive.Entries.Count > options.MaximumPackageEntries)
            throw new InvalidDataException(
                $"Package contains {archive.Entries.Count} entries; limit is {options.MaximumPackageEntries}.");

        var parts = new Dictionary<string, Part>(StringComparer.Ordinal);
        var partNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var budget = new ReadBudget(options.MaximumTotalUncompressedBytes);
        long declaredTotal = 0;
        foreach (var entry in archive.Entries.OrderBy(item => item.FullName, StringComparer.Ordinal))
        {
            var decodedName = ValidateEntryName(entry.FullName, options.MaximumPartUriLength);
            if (string.IsNullOrEmpty(entry.Name)) continue;
            if (!partNames.Add(decodedName))
                throw new InvalidDataException(
                    $"Package contains duplicate part name '{entry.FullName}'.");
            if (entry.Length > options.MaximumPartBytes)
                throw new InvalidDataException(
                    $"Package entry '{entry.FullName}' is {entry.Length} bytes; limit is {options.MaximumPartBytes}.");
            if (entry.CompressedLength == 0 && entry.Length > 0
                || entry.CompressedLength > 0
                    && entry.Length / (double)entry.CompressedLength > options.MaximumCompressionRatio)
                throw new InvalidDataException(
                    $"Package entry '{entry.FullName}' exceeds the {options.MaximumCompressionRatio:R}:1 compression-ratio limit.");
            if (declaredTotal > options.MaximumTotalUncompressedBytes - entry.Length)
                throw new InvalidDataException(
                    $"Package declared uncompressed size exceeds the {options.MaximumTotalUncompressedBytes}-byte aggregate limit.");
            declaredTotal += entry.Length;
            using var entryStream = entry.Open();
            var payload = ReadEntryBytes(
                entryStream,
                entry.FullName,
                options.MaximumPartBytes,
                entry.Length,
                budget);
            parts.Add(entry.FullName, new Part(entry.FullName, payload, TryReadXml(payload)));
        }

        var contentTypes = ReadContentTypes(parts);
        var relationshipData = ReadRelationships(parts);
        var relationshipBindings = ReadRelationshipBindings(parts, relationshipData.ByOwnerAndId);
        var bookmarks = new List<Entity>();
        var revisions = new List<Entity>();
        var annotations = new List<Entity>();
        foreach (var part in parts.Values.OrderBy(item => item.Name, StringComparer.Ordinal))
        {
            if (part.Xml?.Root is null) continue;
            var partUri = PartUri(part.Name);
            if (IsWordStoryPart(part.Name))
            {
                UnidHelper.AssignToAllElementsDeterministic(part.Xml.Root);
                bookmarks.AddRange(ReadBookmarks(part, partUri));
                revisions.AddRange(ReadRevisions(part, partUri));
            }
            if (part.Xml.Root.Name.NamespaceName == AnnotationNamespace
                && part.Xml.Root.Name.LocalName == "annotations")
                annotations.AddRange(ReadAnnotations(part, partUri));
        }

        var media = parts.Values
            .Where(part => IsMediaPart(part.Name))
            .OrderBy(part => part.Name, StringComparer.Ordinal)
            .Select(part =>
            {
                var value = ValueObj(
                    ("contentType", SemanticValue.String(ContentTypeFor(part.Name, contentTypes))),
                    ("size", SemanticValue.Integer(part.Bytes.LongLength)),
                    ("digest", SemanticValue.Digest(
                        "SHA-256", Sha256(part.Bytes), "raw-media-bytes")));
                return new Entity(
                    "media:" + PartUri(part.Name),
                    SemanticChangeFamily.Media,
                    PartUri(part.Name),
                    "media",
                    null,
                    null,
                    value,
                    ValueFingerprint(value));
            })
            .ToArray();

        var registryParts = parts.Values
            .Where(part => IsRegistryPart(part.Name))
            .OrderBy(part => part.Name, StringComparer.Ordinal)
            .Select(part => RegistryEntity(part, relationshipData.ByOwnerAndId))
            .ToArray();

        var opaque = parts.Values
            .Where(IsOpaqueCandidate)
            .OrderBy(part => part.Name, StringComparer.Ordinal)
            .Select(part =>
            {
                var preserveWhitespace = PreserveWhitespaceInOpaqueXml(part.Name);
                var fingerprint = part.Xml is null
                    ? Sha256(part.Bytes)
                    : Sha256(Encoding.UTF8.GetBytes(CanonicalXml(
                        part.Xml,
                        preserveWhitespace,
                        RelationshipAttributeNormalizer(
                            PartUri(part.Name), relationshipData.ByOwnerAndId))));
                var contentType = ContentTypeFor(part.Name, contentTypes);
                var digest = SemanticValue.Digest(
                    "SHA-256",
                    fingerprint,
                    part.Xml is null ? "raw-part-bytes"
                        : preserveWhitespace ? "xml-expanded-names-whitespace-v1"
                        : "xml-expanded-names-v1");
                var value = ValueObj(
                    ("contentType", SemanticValue.String(contentType)),
                    ("size", SemanticValue.Integer(part.Bytes.LongLength)),
                    ("normalizedDigest", digest));
                var identity = ValueObj(
                    ("contentType", SemanticValue.String(contentType)),
                    ("normalizedDigest", digest));
                return new Entity(
                    "opaque:" + PartUri(part.Name),
                    SemanticChangeFamily.OpaquePackagePart,
                    PartUri(part.Name),
                    "package.part",
                    null,
                    null,
                    value,
                    ValueFingerprint(identity));
            })
            .ToArray();

        return new PackageSnapshot(
            relationshipData.Inventory.Cast<Entity>().ToArray(),
            relationshipBindings,
            media,
            bookmarks,
            revisions,
            annotations,
            registryParts,
            opaque);
    }

    internal static byte[] ReadEntryBytes(
        Stream input,
        string entryName,
        long maximumBytes,
        long declaredLength) => ReadEntryBytes(
            input,
            entryName,
            maximumBytes,
            declaredLength,
            aggregateBudget: null);

    private static byte[] ReadEntryBytes(
        Stream input,
        string entryName,
        long maximumBytes,
        long declaredLength,
        ReadBudget? aggregateBudget)
    {
        // entry.Length is controlled by the ZIP central directory. Treat it as an early rejection
        // hint only; a forged value must not turn CopyTo into an unbounded decompression/allocation.
        using var copy = new MemoryStream((int)Math.Min(Math.Max(declaredLength, 0), 81920));
        var buffer = new byte[81920];
        long total = 0;
        int read;
        while ((read = input.Read(buffer, 0, buffer.Length)) > 0)
        {
            if (total > maximumBytes - read)
                throw new InvalidDataException(
                    $"Package entry '{entryName}' exceeds the {maximumBytes}-byte decompressed limit.");
            aggregateBudget?.Consume(read, entryName);
            copy.Write(buffer, 0, read);
            total += read;
        }
        return copy.ToArray();
    }

    private static string ValidateEntryName(string name, int maximumUriLength)
    {
        if (string.IsNullOrEmpty(name))
            throw new InvalidDataException("Package contains an empty ZIP entry name.");
        if (name.StartsWith("/", StringComparison.Ordinal)
            || name.Contains('\\', StringComparison.Ordinal)
            || name.Split('/').Any(segment => segment is "." or ".."))
            throw new InvalidDataException($"Package entry name '{name}' is not a safe OPC part name.");

        for (int index = 0; index < name.Length; index++)
        {
            if (name[index] != '%') continue;
            if (index + 2 >= name.Length || !IsHex(name[index + 1]) || !IsHex(name[index + 2]))
                throw new InvalidDataException($"Package entry name '{name}' has invalid percent encoding.");
            index += 2;
        }

        string decoded;
        try
        {
            decoded = Uri.UnescapeDataString(name);
        }
        catch (UriFormatException exception)
        {
            throw new InvalidDataException($"Package entry name '{name}' is not a valid URI.", exception);
        }
        if (decoded.Length > maximumUriLength)
            throw new InvalidDataException(
                $"Decoded package entry name '{name}' is {decoded.Length} characters; limit is {maximumUriLength}.");
        var decodedSegments = decoded.Split('/');
        bool isDirectory = decoded.EndsWith("/", StringComparison.Ordinal);
        if (decoded.StartsWith("/", StringComparison.Ordinal)
            || decoded.Contains('\\', StringComparison.Ordinal)
            || decoded.Any(char.IsControl)
            || decodedSegments.Where((segment, index) =>
                segment is "." or ".."
                || segment.Length == 0 && index != decodedSegments.Length - 1).Any()
            || !isDirectory && decodedSegments[^1].Length == 0)
            throw new InvalidDataException(
                $"Decoded package entry name '{name}' is not a safe OPC part name.");
        return decoded;
    }

    private static bool IsHex(char value) =>
        value is >= '0' and <= '9'
        || value is >= 'a' and <= 'f'
        || value is >= 'A' and <= 'F';

    private static ContentTypeMap ReadContentTypes(IReadOnlyDictionary<string, Part> parts)
    {
        var defaults = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var overrides = new Dictionary<string, string>(StringComparer.Ordinal);
        if (!parts.TryGetValue("[Content_Types].xml", out var contentTypes)
            || contentTypes.Xml?.Root is null)
            return new ContentTypeMap(defaults, overrides);

        foreach (var element in contentTypes.Xml.Root.Elements())
        {
            var contentType = Attr(element, "ContentType");
            if (string.IsNullOrWhiteSpace(contentType)) continue;
            if (element.Name.LocalName == "Default" && Attr(element, "Extension") is { } extension)
                defaults[extension.TrimStart('.')] = contentType;
            else if (element.Name.LocalName == "Override" && Attr(element, "PartName") is { } partName)
                overrides[PartUri(partName)] = contentType;
        }
        return new ContentTypeMap(defaults, overrides);
    }

    private static RelationshipReadResult ReadRelationships(
        IReadOnlyDictionary<string, Part> parts)
    {
        var entities = new List<RelationshipEntity>();
        var definitions = new Dictionary<(string Owner, string Id), RelationshipInfo>();
        foreach (var part in parts.Values
            .Where(item => item.Name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
            .OrderBy(item => item.Name, StringComparer.Ordinal))
        {
            if (part.Xml?.Root is null) continue;
            var owner = RelationshipOwner(part.Name);
            var grouped = part.Xml.Root.Elements()
                .Where(element => element.Name.LocalName == "Relationship"
                    && (element.Name.NamespaceName == RelationshipNamespace
                        || element.Name.NamespaceName == StrictRelationshipNamespace
                        || string.IsNullOrEmpty(element.Name.NamespaceName)))
                .Select(element => new
                {
                    Id = Attr(element, "Id"),
                    Type = Attr(element, "Type") ?? string.Empty,
                    Target = Attr(element, "Target") ?? string.Empty,
                    Mode = Attr(element, "TargetMode") ?? "Internal",
                })
                .Select(item => new
                {
                    item.Id,
                    item.Type,
                    Target = NormalizeRelationshipTarget(owner, item.Target, item.Mode),
                    item.Mode,
                })
                .OrderBy(item => item.Id, StringComparer.Ordinal)
                .ThenBy(item => item.Type, StringComparer.Ordinal)
                .ThenBy(item => item.Target, StringComparer.Ordinal)
                .ThenBy(item => item.Mode, StringComparer.Ordinal);
            foreach (var relationship in grouped)
            {
                if (string.IsNullOrEmpty(relationship.Id))
                    throw new InvalidDataException(
                        $"Relationship part '{part.Name}' contains a relationship without an Id.");
                var info = new RelationshipInfo(
                    owner,
                    relationship.Id,
                    relationship.Type,
                    relationship.Target,
                    relationship.Mode);
                if (!definitions.TryAdd((owner, relationship.Id), info))
                    throw new InvalidDataException(
                        $"Relationship part '{part.Name}' contains duplicate Id '{relationship.Id}'.");
                var fingerprint = RelationshipFingerprint(info);
                entities.Add(new RelationshipEntity(
                    $"relationship:{owner}:{relationship.Id}",
                    SemanticChangeFamily.Relationship,
                    owner,
                    "relationship",
                    null,
                    ScopeForPart(owner),
                    RelationshipValue(info),
                    fingerprint,
                    relationship.Id));
            }
        }
        return new RelationshipReadResult(entities, definitions);
    }

    private static IReadOnlyList<Entity> ReadRelationshipBindings(
        IReadOnlyDictionary<string, Part> parts,
        IReadOnlyDictionary<(string Owner, string Id), RelationshipInfo> definitions)
    {
        var entities = new List<Entity>();
        foreach (var part in parts.Values
            .Where(item => item.Xml?.Root is not null
                && !item.Name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
            .OrderBy(item => item.Name, StringComparer.Ordinal))
        {
            var owner = PartUri(part.Name);
            var root = part.Xml!.Root!;
            foreach (var attribute in root.DescendantsAndSelf()
                .SelectMany(element => element.Attributes())
                .Where(IsOfficeRelationshipAttribute))
            {
                var element = attribute.Parent!;
                var elementPath = ElementPath(root, element);
                var attributeName = ExpandedName(attribute.Name);
                var key = $"relationship-binding:{owner}:{elementPath}:{attributeName}";
                var anchor = NearestAnchor(element, part.Name);
                RelationshipInfo? relationship = null;
                definitions.TryGetValue((owner, attribute.Value), out relationship);
                var value = relationship is null
                    ? ValueObj(
                        ("resolved", SemanticValue.Boolean(false)),
                        ("reference", SemanticValue.String(attribute.Value)),
                        ("attribute", SemanticValue.String(attributeName)))
                    : ValueObj(
                        ("resolved", SemanticValue.Boolean(true)),
                        ("attribute", SemanticValue.String(attributeName)),
                        ("type", SemanticValue.String(relationship.Type)),
                        ("target", SemanticValue.String(relationship.Target)),
                        ("targetMode", SemanticValue.String(relationship.Mode)));
                entities.Add(new Entity(
                    key,
                    SemanticChangeFamily.Relationship,
                    owner,
                    $"relationship.binding[{elementPath}]",
                    anchor,
                    ScopeForPart(owner),
                    value,
                    ValueFingerprint(value),
                    key));
            }
        }
        return entities;
    }

    private static bool IsOfficeRelationshipAttribute(XAttribute attribute) =>
        attribute.Name.NamespaceName == OfficeRelationshipNamespace
        || attribute.Name.NamespaceName == StrictOfficeRelationshipNamespace;

    private static SemanticValue RelationshipValue(RelationshipInfo relationship) => ValueObj(
        ("type", SemanticValue.String(relationship.Type)),
        ("target", SemanticValue.String(relationship.Target)),
        ("targetMode", SemanticValue.String(relationship.Mode)));

    private static string RelationshipFingerprint(RelationshipInfo relationship) => string.Join(
        "\u001f",
        relationship.Owner,
        relationship.Type,
        relationship.Target,
        relationship.Mode);

    internal static string NormalizeRelationshipTarget(string owner, string target, string mode)
    {
        // Open XML SDK package cloning commonly rewrites internal targets from owner-relative
        // ("styles.xml") to package-absolute ("/word/styles.xml"). They identify the same OPC
        // part and therefore must not become delete+insert relationship noise. External targets
        // remain byte-for-byte meaningful (apart from relationship-id churn handled above).
        if (string.Equals(mode, "External", StringComparison.OrdinalIgnoreCase)
            || string.IsNullOrEmpty(target))
            return target;

        try
        {
            var baseUri = new Uri("http://docxodus.invalid" +
                (owner == "/" ? "/" : owner), UriKind.Absolute);
            var resolved = new Uri(baseUri, target);
            return resolved.PathAndQuery + resolved.Fragment;
        }
        catch (UriFormatException)
        {
            // #456 owns package-validity findings. This fallback remains deterministic and
            // preserves a malformed target verbatim so it is never silently discarded.
            return target;
        }
    }

    private static IEnumerable<Entity> ReadBookmarks(Part part, string partUri)
    {
        var root = part.Xml!.Root!;
        var endsById = root.Descendants()
            .Where(element => element.Name.LocalName == "bookmarkEnd"
                && IsWordNamespace(element.Name.NamespaceName))
            .Where(element => Attr(element, "id") is not null)
            .GroupBy(element => Attr(element, "id")!, StringComparer.Ordinal)
            .ToDictionary(
                group => group.Key,
                group => new Queue<XElement>(group),
                StringComparer.Ordinal);
        var starts = part.Xml!.Descendants()
            .Where(element => element.Name.LocalName == "bookmarkStart"
                && IsWordNamespace(element.Name.NamespaceName))
            .ToArray();
        var grouped = starts.GroupBy(element => Attr(element, "name") ?? string.Empty)
            .OrderBy(group => group.Key, StringComparer.Ordinal);
        foreach (var group in grouped)
        {
            int ordinal = 0;
            foreach (var bookmark in group)
            {
                var anchor = NearestAnchor(bookmark, part.Name);
                var name = Attr(bookmark, "name") ?? string.Empty;
                var nativeId = Attr(bookmark, "id");
                XElement? end = null;
                if (nativeId is not null && endsById.TryGetValue(nativeId, out var candidates)
                    && candidates.Count > 0)
                    end = candidates.Dequeue();
                var endAnchor = end is null ? null : NearestAnchor(end, part.Name);
                var startPath = ElementPath(root, bookmark);
                var endPath = end is null ? null : ElementPath(root, end);
                var value = ValueObj(
                    ("name", SemanticValue.String(name)),
                    ("columnFirst", SemanticValue.Integer(ParseLong(Attr(bookmark, "colFirst")))),
                    ("columnLast", SemanticValue.Integer(ParseLong(Attr(bookmark, "colLast")))),
                    ("startAnchor", SemanticValue.String(anchor)),
                    ("endAnchor", SemanticValue.String(endAnchor)),
                    ("startPath", SemanticValue.String(startPath)),
                    ("endPath", SemanticValue.String(endPath)));
                var fingerprint = ValueFingerprint(ValueObj(
                    ("name", SemanticValue.String(name)),
                    ("columnFirst", SemanticValue.Integer(ParseLong(Attr(bookmark, "colFirst")))),
                    ("columnLast", SemanticValue.Integer(ParseLong(Attr(bookmark, "colLast"))))));
                yield return new Entity(
                    $"bookmark:{partUri}:{name}:{ordinal++}",
                    SemanticChangeFamily.Bookmark,
                    partUri,
                    "bookmark",
                    anchor,
                    ScopeForPart(partUri),
                    value,
                    fingerprint,
                    string.Join("\u001f", startPath, endPath));
            }
        }
    }

    private static IEnumerable<Entity> ReadRevisions(Part part, string partUri)
    {
        var root = part.Xml!.Root!;
        var revisions = root.Descendants()
            .Where(element => RevisionNames.Contains(element.Name.LocalName)
                && IsWordNamespace(element.Name.NamespaceName))
            .ToArray();
        var ordinals = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (var revision in revisions)
        {
            var kind = revision.Name.LocalName;
            var nativeId = Attr(revision, "id");
            var identity = nativeId is null ? kind : kind + ":" + nativeId;
            ordinals.TryGetValue(identity, out var ordinal);
            ordinals[identity] = ordinal + 1;
            var anchor = NearestAnchor(revision, part.Name);
            var structuralPath = ElementPath(root, revision);
            var normalizedRevision = new XElement(revision);
            normalizedRevision.DescendantsAndSelf()
                .Attributes()
                .Where(attribute => attribute.Name.LocalName == "id"
                    && IsWordNamespace(attribute.Name.NamespaceName))
                .Remove();
            var value = ValueObj(
                ("kind", SemanticValue.String(kind)),
                ("author", SemanticValue.String(Attr(revision, "author"))),
                ("date", SemanticValue.String(Attr(revision, "date"))),
                ("text", SemanticValue.String(string.Concat(revision.DescendantsAndSelf()
                    .Where(element => element.Name.LocalName is "t" or "delText" or "instrText" or "delInstrText")
                    .Select(element => element.Value)))),
                ("normalizedDigest", SemanticValue.Digest(
                    "SHA-256",
                    Sha256(Encoding.UTF8.GetBytes(CanonicalXml(normalizedRevision))),
                    "xml-expanded-names-comments-pi-v1")));
            yield return new Entity(
                $"revision:{partUri}:{identity}:{ordinal}",
                SemanticChangeFamily.Revision,
                partUri,
                "revision",
                anchor,
                ScopeForPart(partUri),
                value,
                ValueFingerprint(value),
                structuralPath);
        }
    }

    private static IEnumerable<Entity> ReadAnnotations(Part part, string partUri)
    {
        foreach (var annotation in part.Xml!.Root!.Elements()
            .Where(element => element.Name.NamespaceName == AnnotationNamespace
                && element.Name.LocalName == "annotation")
            .OrderBy(element => Attr(element, "id"), StringComparer.Ordinal))
        {
            var id = Attr(annotation, "id") ?? string.Empty;
            var bookmarkName = annotation.Descendants()
                .FirstOrDefault(element => element.Name.NamespaceName == AnnotationNamespace
                    && element.Name.LocalName == "range")?
                .Attributes().FirstOrDefault(attribute => attribute.Name.LocalName == "bookmarkName")?.Value;
            var canonical = CanonicalXml(annotation);
            var value = ValueObj(
                ("id", SemanticValue.String(id)),
                ("labelId", SemanticValue.String(Attr(annotation, "labelId"))),
                ("label", SemanticValue.String(Attr(annotation, "label"))),
                ("color", SemanticValue.String(Attr(annotation, "color"))),
                ("author", SemanticValue.String(Attr(annotation, "author"))),
                ("created", SemanticValue.String(Attr(annotation, "created"))),
                ("bookmarkName", SemanticValue.String(bookmarkName)),
                ("normalizedDigest", SemanticValue.Digest(
                    "SHA-256",
                    Sha256(Encoding.UTF8.GetBytes(canonical)),
                    "docxodus-annotation-v1")));
            yield return new Entity(
                $"annotation:{partUri}:{id}",
                SemanticChangeFamily.Annotation,
                partUri,
                "annotation",
                null,
                null,
                value,
                canonical,
                $"{partUri}\u001f{id}");
        }
    }

    private static void CompareEntities(
        IReadOnlyList<Entity> left,
        IReadOnlyList<Entity> right,
        List<SemanticChangeDraft> result)
    {
        var unmatchedLeft = new HashSet<int>(Enumerable.Range(0, left.Count));
        var unmatchedRight = new HashSet<int>(Enumerable.Range(0, right.Count));

        // Match semantic equals before native keys. This is what makes a coordinated rId rewrite
        // serialization-only even when another relationship takes over the old rId.
        MatchEntityPairs(
            left,
            right,
            unmatchedLeft,
            unmatchedRight,
            EntityExactKey,
            (_, _) => { });

        MatchEntityPairs(
            left,
            right,
            unmatchedLeft,
            unmatchedRight,
            entity => entity.Key,
            (before, after) => EmitEntityPair(before, after, result));

        // A semantically identical item at a new structural location is a move when identity is
        // otherwise recoverable (bookmarks/revisions are the principal callers).
        MatchEntityPairs(
            left,
            right,
            unmatchedLeft,
            unmatchedRight,
            EntityMeaningKey,
            (before, after) => EmitEntityPair(before, after, result));

        MatchEntityPairs(
            left,
            right,
            unmatchedLeft,
            unmatchedRight,
            EntityLocationKey,
            (before, after) => EmitEntityPair(before, after, result),
            requireNonEmptyKey: true);

        // Pair an unambiguous final residue within a family/part/path as a modification rather
        // than manufacturing a delete+insert when a native identity changed with the content.
        foreach (var groupKey in unmatchedLeft
            .Select(index => EntityGroupKey(left[index]))
            .Concat(unmatchedRight.Select(index => EntityGroupKey(right[index])))
            .Distinct(StringComparer.Ordinal)
            .OrderBy(key => key, StringComparer.Ordinal))
        {
            var leftResidue = unmatchedLeft
                .Where(index => EntityGroupKey(left[index]) == groupKey)
                .ToArray();
            var rightResidue = unmatchedRight
                .Where(index => EntityGroupKey(right[index]) == groupKey)
                .ToArray();
            if (leftResidue.Length != 1 || rightResidue.Length != 1) continue;
            EmitEntityPair(left[leftResidue[0]], right[rightResidue[0]], result);
            unmatchedLeft.Remove(leftResidue[0]);
            unmatchedRight.Remove(rightResidue[0]);
        }

        foreach (int index in unmatchedLeft.OrderBy(index => left[index].Key, StringComparer.Ordinal))
            EmitEntityPair(left[index], null, result);
        foreach (int index in unmatchedRight.OrderBy(index => right[index].Key, StringComparer.Ordinal))
            EmitEntityPair(null, right[index], result);
    }

    private static void MatchEntityPairs(
        IReadOnlyList<Entity> left,
        IReadOnlyList<Entity> right,
        HashSet<int> unmatchedLeft,
        HashSet<int> unmatchedRight,
        Func<Entity, string> keySelector,
        Action<Entity, Entity> onMatch,
        bool requireNonEmptyKey = false)
    {
        var rightByKey = unmatchedRight
            .GroupBy(index => keySelector(right[index]), StringComparer.Ordinal)
            .ToDictionary(
                group => group.Key,
                group => new Queue<int>(group.OrderBy(index => right[index].Key, StringComparer.Ordinal)),
                StringComparer.Ordinal);
        foreach (int leftIndex in unmatchedLeft
            .OrderBy(index => left[index].Key, StringComparer.Ordinal)
            .ToArray())
        {
            string key = keySelector(left[leftIndex]);
            if (requireNonEmptyKey && string.IsNullOrEmpty(key)) continue;
            if (!rightByKey.TryGetValue(key, out var candidates)) continue;
            while (candidates.Count > 0 && !unmatchedRight.Contains(candidates.Peek()))
                candidates.Dequeue();
            if (candidates.Count == 0) continue;
            int rightIndex = candidates.Dequeue();
            onMatch(left[leftIndex], right[rightIndex]);
            unmatchedLeft.Remove(leftIndex);
            unmatchedRight.Remove(rightIndex);
        }
    }

    private static void EmitEntityPair(
        Entity? before,
        Entity? after,
        List<SemanticChangeDraft> result)
    {
        if (before is not null && after is not null
            && EntityExactKey(before) == EntityExactKey(after))
            return;
        var exemplar = after ?? before!;
        bool isMove = before is not null && after is not null
            && EntityMeaningKey(before) == EntityMeaningKey(after)
            && EntityLocationKey(before) != EntityLocationKey(after);
        var operation = before is null ? SemanticChangeOperation.Insert
            : after is null ? SemanticChangeOperation.Delete
            : isMove ? SemanticChangeOperation.Move
            : SemanticChangeOperation.Modify;
        result.Add(new SemanticChangeDraft(
            operation,
            exemplar.Family,
            exemplar.PartUri,
            exemplar.Path,
            before?.Anchor,
            after?.Anchor,
            before?.Scope,
            after?.Scope,
            isMove ? $"package:{(int)exemplar.Family}:{before!.Key}:{after!.Key}" : null,
            before?.Value ?? SemanticValue.Absent,
            after?.Value ?? SemanticValue.Absent));
    }

    private static string EntityExactKey(Entity entity) => string.Join(
        "\u001e", EntityMeaningKey(entity), EntityLocationKey(entity));

    private static string EntityMeaningKey(Entity entity) => string.Join(
        "\u001e", EntityGroupKey(entity), entity.Fingerprint);

    private static string EntityLocationKey(Entity entity)
    {
        if (entity.LocationKey is not null) return entity.LocationKey;
        if (entity.Scope is null && entity.Anchor is null) return string.Empty;
        return string.Join("\u001f", entity.Scope ?? string.Empty, entity.Anchor ?? string.Empty);
    }

    private static string EntityGroupKey(Entity entity) => string.Join(
        "\u001f",
        ((int)entity.Family).ToString(System.Globalization.CultureInfo.InvariantCulture),
        entity.PartUri,
        entity.Path);

    private static XDocument? TryReadXml(byte[] bytes)
    {
        if (bytes.Length == 0) return null;
        try
        {
            using var stream = new MemoryStream(bytes, writable: false);
            using var reader = XmlReader.Create(stream, new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                MaxCharactersInDocument = Math.Max(bytes.LongLength * 4, 1024),
            });
            return XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        }
        catch (XmlException)
        {
            return null;
        }
    }

    private static string CanonicalXml(
        XElement element,
        bool preserveWhitespace = false,
        Func<XAttribute, string>? attributeNormalizer = null)
    {
        var builder = new StringBuilder();
        WriteCanonical(element, builder, preserveWhitespace, attributeNormalizer);
        return builder.ToString();
    }

    private static string CanonicalXml(
        XDocument document,
        bool preserveWhitespace,
        Func<XAttribute, string>? attributeNormalizer = null)
    {
        var builder = new StringBuilder();
        foreach (var node in document.Nodes())
            WriteCanonicalNode(node, builder, preserveWhitespace, attributeNormalizer, parent: null);
        return builder.ToString();
    }

    private static void WriteCanonical(
        XElement element,
        StringBuilder builder,
        bool preserveWhitespace,
        Func<XAttribute, string>? attributeNormalizer)
    {
        builder.Append('<').Append('{').Append(element.Name.NamespaceName).Append('}')
            .Append(element.Name.LocalName);
        foreach (var attribute in element.Attributes()
            .Where(attribute => !attribute.IsNamespaceDeclaration && attribute.Name != PtOpenXml.Unid)
            .OrderBy(attribute => attribute.Name.NamespaceName, StringComparer.Ordinal)
            .ThenBy(attribute => attribute.Name.LocalName, StringComparer.Ordinal))
        {
            var value = attributeNormalizer?.Invoke(attribute) ?? attribute.Value;
            builder.Append(" a{").Append(attribute.Name.NamespaceName).Append('}')
                .Append(attribute.Name.LocalName).Append('=')
                .Append(Convert.ToBase64String(Encoding.UTF8.GetBytes(value)));
        }
        builder.Append('>');
        foreach (var node in element.Nodes())
            WriteCanonicalNode(node, builder, preserveWhitespace, attributeNormalizer, element);
        builder.Append("</>");
    }

    private static void WriteCanonicalNode(
        XNode node,
        StringBuilder builder,
        bool preserveWhitespace,
        Func<XAttribute, string>? attributeNormalizer,
        XElement? parent)
    {
        switch (node)
        {
            case XElement child:
                WriteCanonical(child, builder, preserveWhitespace, attributeNormalizer);
                break;
            case XCData cdata:
                builder.Append(" c=").Append(
                    Convert.ToBase64String(Encoding.UTF8.GetBytes(cdata.Value)));
                break;
            case XText text when parent is not null && (preserveWhitespace
                || !string.IsNullOrWhiteSpace(text.Value)
                || parent.AncestorsAndSelf().Any(ancestor =>
                    (string?)ancestor.Attribute(XNamespace.Xml + "space") == "preserve")):
                builder.Append(" t=").Append(
                    Convert.ToBase64String(Encoding.UTF8.GetBytes(text.Value)));
                break;
            case XComment comment:
                builder.Append(" m=").Append(
                    Convert.ToBase64String(Encoding.UTF8.GetBytes(comment.Value)));
                break;
            case XProcessingInstruction instruction:
                builder.Append(" i=")
                    .Append(Convert.ToBase64String(Encoding.UTF8.GetBytes(instruction.Target)))
                    .Append(':')
                    .Append(Convert.ToBase64String(Encoding.UTF8.GetBytes(instruction.Data)));
                break;
        }
    }

    private static Func<XAttribute, string> RelationshipAttributeNormalizer(
        string owner,
        IReadOnlyDictionary<(string Owner, string Id), RelationshipInfo> definitions) => attribute =>
    {
        if (!IsOfficeRelationshipAttribute(attribute)) return attribute.Value;
        return definitions.TryGetValue((owner, attribute.Value), out var relationship)
            ? "semantic-relationship:" + RelationshipFingerprint(relationship)
            : "unresolved-relationship:" + attribute.Value;
    };

    private static Entity RegistryEntity(
        Part part,
        IReadOnlyDictionary<(string Owner, string Id), RelationshipInfo> relationships)
    {
        var partUri = PartUri(part.Name);
        var family = part.Name == "word/numbering.xml"
            ? SemanticChangeFamily.Numbering
            : SemanticChangeFamily.Style;
        var path = part.Name == "word/numbering.xml"
            ? "numbering.registry.package"
            : part.Name.StartsWith("word/theme/", StringComparison.Ordinal)
                ? "theme.registry.package"
                : "style.registry.package";
        string fingerprint = part.Xml is null
            ? Sha256(part.Bytes)
            : Sha256(Encoding.UTF8.GetBytes(CanonicalXml(
                part.Xml,
                preserveWhitespace: false,
                RelationshipAttributeNormalizer(partUri, relationships))));
        var value = ValueObj(
            ("size", SemanticValue.Integer(part.Bytes.LongLength)),
            ("normalizedDigest", SemanticValue.Digest(
                "SHA-256",
                fingerprint,
                part.Xml is null ? "raw-part-bytes" : "xml-expanded-names-comments-pi-v1")));
        return new Entity(
            "registry:" + partUri,
            family,
            partUri,
            path,
            null,
            null,
            value,
            ValueFingerprint(value));
    }

    private static bool PreserveWhitespaceInOpaqueXml(string name) =>
        // Word-owned XML parts are declarative element/attribute vocabularies where indentation is
        // serialization. Unknown/vendor parts, especially customXml, may use whitespace as data and
        // therefore receive a whitespace-preserving fingerprint.
        name is not ("word/settings.xml" or "word/webSettings.xml" or "word/fontTable.xml")
        && !name.StartsWith("docProps/", StringComparison.Ordinal);

    private static bool IsOpaqueCandidate(Part part)
    {
        var name = part.Name;
        if (name == "[Content_Types].xml" || name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
            return false;
        if (IsMediaPart(name) || IsWordStoryPart(name)) return false;
        if (IsRegistryPart(name)) return false;
        if (part.Xml?.Root?.Name is { NamespaceName: AnnotationNamespace, LocalName: "annotations" })
            return false;
        return true;
    }

    private static bool IsRegistryPart(string name) =>
        name is "word/styles.xml" or "word/numbering.xml"
        || name.StartsWith("word/theme/", StringComparison.Ordinal);

    private static bool IsWordStoryPart(string name) =>
        name == "word/document.xml"
        || name == "word/footnotes.xml"
        || name == "word/endnotes.xml"
        || name == "word/comments.xml"
        || (name.StartsWith("word/header", StringComparison.Ordinal) && name.EndsWith(".xml", StringComparison.Ordinal))
        || (name.StartsWith("word/footer", StringComparison.Ordinal) && name.EndsWith(".xml", StringComparison.Ordinal));

    private static bool IsMediaPart(string name) =>
        name.StartsWith("media/", StringComparison.Ordinal)
        || name.Contains("/media/", StringComparison.Ordinal)
        || name.StartsWith("word/embeddings/", StringComparison.Ordinal);

    private static bool IsWordNamespace(string value) =>
        value == WordNamespace || value == StrictWordNamespace;

    private static string RelationshipOwner(string relationshipPart)
    {
        if (relationshipPart == "_rels/.rels") return "/";
        int marker = relationshipPart.LastIndexOf("/_rels/", StringComparison.Ordinal);
        if (marker < 0) return PartUri(relationshipPart);
        var prefix = relationshipPart.Substring(0, marker + 1);
        var file = relationshipPart.Substring(marker + "/_rels/".Length);
        if (file.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
            file = file.Substring(0, file.Length - ".rels".Length);
        return PartUri(prefix + file);
    }

    private static string ElementPath(XElement root, XElement element)
    {
        var segments = new Stack<string>();
        XElement? current = element;
        while (current is not null)
        {
            int ordinal = current.ElementsBeforeSelf()
                .Count(sibling => sibling.Name == current.Name) + 1;
            segments.Push($"{ExpandedName(current.Name)}[{ordinal}]");
            if (current == root) break;
            current = current.Parent;
        }
        return "/" + string.Join("/", segments);
    }

    private static string ExpandedName(XName name) =>
        $"{{{name.NamespaceName}}}{name.LocalName}";

    private static string? NearestAnchor(XElement element, string entryName)
    {
        foreach (var candidate in element.AncestorsAndSelf())
        {
            string? kind = candidate.Name.LocalName switch
            {
                "p" => "p",
                "tbl" => "tbl",
                "tr" => "tr",
                "tc" => "tc",
                "sdt" => "sdt",
                "sectPr" => "sec",
                _ => null,
            };
            if (kind is null) continue;
            var unid = (string?)candidate.Attribute(PtOpenXml.Unid);
            if (string.IsNullOrWhiteSpace(unid)) continue;
            return $"{kind}:{ScopeForPart(PartUri(entryName))}:{unid}";
        }
        return null;
    }

    private static string? ScopeForPart(string partUri)
    {
        var name = partUri.TrimStart('/');
        if (name == "word/document.xml") return "body";
        if (name == "word/footnotes.xml") return "fn";
        if (name == "word/endnotes.xml") return "en";
        if (name == "word/comments.xml") return "cmt";
        if (name.StartsWith("word/header", StringComparison.Ordinal))
            return "hdr" + Digits(Path.GetFileNameWithoutExtension(name));
        if (name.StartsWith("word/footer", StringComparison.Ordinal))
            return "ftr" + Digits(Path.GetFileNameWithoutExtension(name));
        return null;
    }

    private static string Digits(string value) => new(value.Where(char.IsDigit).ToArray());

    private static string? Attr(XElement element, string localName) =>
        element.Attributes().FirstOrDefault(attribute => attribute.Name.LocalName == localName)?.Value;

    private static string PartUri(string entryName) =>
        entryName.StartsWith("/", StringComparison.Ordinal) ? entryName : "/" + entryName;

    private static string Sha256(byte[] bytes) =>
        Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

    private static long? ParseLong(string? value) =>
        long.TryParse(value, System.Globalization.NumberStyles.Integer,
            System.Globalization.CultureInfo.InvariantCulture, out var parsed) ? parsed : null;

    private static SemanticValue ValueObj(params (string Name, SemanticValue Value)[] properties) =>
        SemanticValue.Object(properties.Select(property =>
            new SemanticProperty(property.Name, property.Value)));

    private static string ValueFingerprint(SemanticValue value)
    {
        var wrapper = new SemanticChangeSet(new[]
        {
            new SemanticChange
            {
                Id = "fingerprint",
                Operation = SemanticChangeOperation.Modify,
                Family = SemanticChangeFamily.OpaquePackagePart,
                PartUri = "/",
                Path = "fingerprint",
                Before = value,
                After = SemanticValue.Absent,
            },
        });
        return wrapper.ToJson(indented: false);
    }

    private static string ContentTypeFor(string name, ContentTypeMap contentTypes)
    {
        if (contentTypes.Overrides.TryGetValue(PartUri(name), out var overridden))
            return overridden;
        var extension = Path.GetExtension(name).ToLowerInvariant();
        if (contentTypes.Defaults.TryGetValue(extension.TrimStart('.'), out var declared))
            return declared;
        return extension switch
        {
            ".xml" => "application/xml",
            ".rels" => "application/vnd.openxmlformats-package.relationships+xml",
            ".png" => "image/png",
            ".jpg" or ".jpeg" => "image/jpeg",
            ".gif" => "image/gif",
            ".bmp" => "image/bmp",
            ".tif" or ".tiff" => "image/tiff",
            ".svg" => "image/svg+xml",
            _ => "application/octet-stream",
        };
    }

    private sealed record Part(string Name, byte[] Bytes, XDocument? Xml);

    private sealed record ContentTypeMap(
        IReadOnlyDictionary<string, string> Defaults,
        IReadOnlyDictionary<string, string> Overrides);

    private sealed class ReadBudget
    {
        private long _remaining;

        public ReadBudget(long maximumBytes) => _remaining = maximumBytes;

        public void Consume(int count, string entryName)
        {
            if (_remaining < count)
                throw new InvalidDataException(
                    $"Package exceeds the aggregate decompressed-byte limit while reading '{entryName}'.");
            _remaining -= count;
        }
    }

    private sealed record RelationshipInfo(
        string Owner,
        string Id,
        string Type,
        string Target,
        string Mode);

    private sealed record RelationshipReadResult(
        IReadOnlyList<RelationshipEntity> Inventory,
        IReadOnlyDictionary<(string Owner, string Id), RelationshipInfo> ByOwnerAndId);

    private record Entity(
        string Key,
        SemanticChangeFamily Family,
        string PartUri,
        string Path,
        string? Anchor,
        string? Scope,
        SemanticValue Value,
        string Fingerprint,
        string? LocationKey = null);

    private sealed record RelationshipEntity(
        string Key,
        SemanticChangeFamily Family,
        string PartUri,
        string Path,
        string? Anchor,
        string? Scope,
        SemanticValue Value,
        string Fingerprint,
        string? RelationshipId,
        string? LocationKey = null)
        : Entity(Key, Family, PartUri, Path, Anchor, Scope, Value, Fingerprint, LocationKey);

    private sealed record PackageSnapshot(
        IReadOnlyList<Entity> Relationships,
        IReadOnlyList<Entity> RelationshipBindings,
        IReadOnlyList<Entity> Media,
        IReadOnlyList<Entity> Bookmarks,
        IReadOnlyList<Entity> Revisions,
        IReadOnlyList<Entity> Annotations,
        IReadOnlyList<Entity> RegistryParts,
        IReadOnlyList<Entity> OpaqueParts);
}
