// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace Docxodus.Verification;

/// <summary>
/// Projects bounded, same-pass package-manifest inspection into semantic facts not represented by
/// the IR. It is intentionally hidden behind <see cref="ISemanticPackageChangeDetector"/> so the
/// public semantic-change schema remains independent of package inspection details.
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
        ArgumentNullException.ThrowIfNull(options.PackageOptions);
        if (!options.IncludePackageChanges)
        {
            var leftManifest = PackageManifestGenerator.Generate(leftBytes, options.PackageOptions);
            var rightManifest = PackageManifestGenerator.Generate(rightBytes, options.PackageOptions);
            EnsureValid(leftManifest, "left");
            EnsureValid(rightManifest, "right");
            return Array.Empty<SemanticChangeDraft>();
        }

        var leftInspection = PackageManifestGenerator.Inspect(leftBytes, options.PackageOptions);
        var rightInspection = PackageManifestGenerator.Inspect(rightBytes, options.PackageOptions);
        EnsureValid(leftInspection.Manifest, "left");
        EnsureValid(rightInspection.Manifest, "right");
        var left = Read(leftInspection);
        var right = Read(rightInspection);
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

    private static void EnsureValid(PackageManifest manifest, string side)
    {
        if (manifest.IsValid) return;
        var errors = manifest.Findings
            .Where(finding => finding.Severity == VerificationFindingSeverity.Error)
            .Select(finding => finding.Code + FormatLocation(finding.Location))
            .ToArray();
        throw new InvalidDataException(
            $"The {side} package failed manifest preflight: {string.Join(", ", errors)}.");
    }

    private static string FormatLocation(ChangeLocation? location)
    {
        if (location is null) return string.Empty;
        var value = location.EntryUri
            ?? location.OwnerUri
            ?? location.TargetUri
            ?? location.PropertyPath;
        return value is null ? string.Empty : $" ({value})";
    }

    private static PackageSnapshot Read(PackageManifestInspection inspection)
    {
        var parts = new Dictionary<string, Part>(StringComparer.Ordinal);
        foreach (var inspected in inspection.Entries)
        {
            var entry = inspected.ManifestEntry;
            if (entry.Uri.EndsWith("/", StringComparison.Ordinal)) continue;
            if (entry.Occurrence != 0 || !inspected.PayloadWasRead)
                throw new InvalidDataException(
                    $"Valid manifest inspection is incomplete for '{entry.Uri}'.");
            var name = entry.Uri.TrimStart('/');
            parts.Add(name, new Part(
                name,
                inspected.Xml,
                entry.ContentType,
                entry.Size,
                entry.RawBytesDigest));
        }

        var relationshipData = ReadRelationships(inspection.Manifest.Relationships);
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
                    ("contentType", SemanticValue.String(ContentTypeFor(part))),
                    ("size", SemanticValue.Integer(part.Size)),
                    ("digest", SemanticValue.Digest(
                        part.RawBytesDigest!.Algorithm,
                        part.RawBytesDigest.Value,
                        "raw-media-bytes")));
                return new Entity(
                    "media:" + PartUri(part.Name),
                    SemanticChangeFamily.Media,
                    new ChangeLocation
                    {
                        EntryUri = PartUri(part.Name),
                        PropertyPath = "media",
                    },
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
                    ? part.RawBytesDigest!.Value
                    : XmlSemanticNormalizer.Digest(
                        part.Xml,
                        PartUri(part.Name),
                        ignoreFormattingWhitespace: !preserveWhitespace,
                        includeAttribute: ExcludeGeneratedUnid,
                        attributeValueNormalizer: RelationshipAttributeNormalizer(
                            PartUri(part.Name), relationshipData.ByOwnerAndId)).Value;
                var contentType = ContentTypeFor(part);
                var digest = SemanticValue.Digest(
                    "SHA-256",
                    fingerprint,
                    part.Xml is null ? "raw-part-bytes"
                        : preserveWhitespace ? "xml-expanded-names-whitespace-v1"
                        : "xml-expanded-names-v1");
                var value = part.Xml is null
                    ? ValueObj(
                        ("contentType", SemanticValue.String(contentType)),
                        ("size", SemanticValue.Integer(part.Size)),
                        ("normalizedDigest", digest))
                    : ValueObj(
                        ("contentType", SemanticValue.String(contentType)),
                        ("normalizedDigest", digest));
                var identity = ValueObj(
                    ("contentType", SemanticValue.String(contentType)),
                    ("normalizedDigest", digest));
                return new Entity(
                    "opaque:" + PartUri(part.Name),
                    SemanticChangeFamily.OpaquePackagePart,
                    new ChangeLocation
                    {
                        EntryUri = PartUri(part.Name),
                        PropertyPath = "package.part",
                    },
                    null,
                    null,
                    value,
                    ValueFingerprint(identity));
            })
            .ToArray();

        return new PackageSnapshot(
            relationshipData.Inventory.ToArray(),
            relationshipBindings,
            media,
            bookmarks,
            revisions,
            annotations,
            registryParts,
            opaque);
    }

    private static RelationshipReadResult ReadRelationships(
        IReadOnlyList<PackageRelationship> relationships)
    {
        var entities = new List<Entity>();
        var definitions = new Dictionary<(string Owner, string Id), RelationshipInfo>();
        foreach (var relationship in relationships
            .OrderBy(item => item.OwnerUri, StringComparer.Ordinal)
            .ThenBy(item => item.Id, StringComparer.Ordinal)
            .ThenBy(item => item.Type, StringComparer.Ordinal)
            .ThenBy(item => item.Target, StringComparer.Ordinal))
        {
            var target = relationship.TargetMode == "Internal"
                ? relationship.ResolvedTargetUri ?? relationship.Target
                : relationship.Target;
            var info = new RelationshipInfo(
                relationship.OwnerUri,
                relationship.Id,
                relationship.Type,
                target,
                relationship.TargetMode);
            if (!definitions.TryAdd((relationship.OwnerUri, relationship.Id), info))
                throw new InvalidDataException(
                    $"Valid manifest repeats relationship '{relationship.Id}' for " +
                    $"'{relationship.OwnerUri}'.");
            var fingerprint = RelationshipFingerprint(info);
            entities.Add(new Entity(
                $"relationship:{relationship.OwnerUri}:{relationship.Id}",
                SemanticChangeFamily.Relationship,
                new ChangeLocation
                {
                    EntryUri = relationship.OwnerUri,
                    OwnerUri = relationship.OwnerUri,
                    RelationshipId = relationship.Id,
                    TargetUri = target,
                    PropertyPath = "relationship",
                },
                null,
                ScopeForPart(relationship.OwnerUri),
                RelationshipValue(info),
                fingerprint));
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
                    new ChangeLocation
                    {
                        EntryUri = owner,
                        OwnerUri = owner,
                        RelationshipId = attribute.Value,
                        TargetUri = relationship?.Target,
                        PropertyPath = $"relationship.binding[{elementPath}]",
                    },
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
                    new ChangeLocation
                    {
                        EntryUri = partUri,
                        PropertyPath = "bookmark",
                    },
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
            var normalizedRevision = XmlSemanticNormalizer.Digest(
                revision,
                partUri,
                ignoreFormattingWhitespace: true,
                includeAttribute: IncludeRevisionAttribute);
            var value = ValueObj(
                ("kind", SemanticValue.String(kind)),
                ("author", SemanticValue.String(Attr(revision, "author"))),
                ("date", SemanticValue.String(Attr(revision, "date"))),
                ("text", SemanticValue.String(string.Concat(revision.DescendantsAndSelf()
                    .Where(element => element.Name.LocalName is "t" or "delText" or "instrText" or "delInstrText")
                    .Select(element => element.Value)))),
                ("normalizedDigest", SemanticValue.Digest(
                    normalizedRevision.Algorithm,
                    normalizedRevision.Value,
                    "xml-expanded-names-comments-pi-v1")));
            yield return new Entity(
                $"revision:{partUri}:{identity}:{ordinal}",
                SemanticChangeFamily.Revision,
                new ChangeLocation
                {
                    EntryUri = partUri,
                    PropertyPath = "revision",
                },
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
            var normalized = XmlSemanticNormalizer.Digest(
                annotation,
                partUri,
                ignoreFormattingWhitespace: false,
                includeAttribute: ExcludeGeneratedUnid);
            var value = ValueObj(
                ("id", SemanticValue.String(id)),
                ("labelId", SemanticValue.String(Attr(annotation, "labelId"))),
                ("label", SemanticValue.String(Attr(annotation, "label"))),
                ("color", SemanticValue.String(Attr(annotation, "color"))),
                ("author", SemanticValue.String(Attr(annotation, "author"))),
                ("created", SemanticValue.String(Attr(annotation, "created"))),
                ("bookmarkName", SemanticValue.String(bookmarkName)),
                ("normalizedDigest", SemanticValue.Digest(
                    normalized.Algorithm,
                    normalized.Value,
                    "docxodus-annotation-v1")));
            yield return new Entity(
                $"annotation:{partUri}:{id}",
                SemanticChangeFamily.Annotation,
                new ChangeLocation
                {
                    EntryUri = partUri,
                    PropertyPath = "annotation",
                },
                null,
                null,
                value,
                normalized.Value,
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
            exemplar.Location.EntryUri ?? exemplar.Location.OwnerUri ?? "/",
            exemplar.Location.PropertyPath ?? "package",
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
        entity.Location.EntryUri ?? entity.Location.OwnerUri ?? string.Empty,
        entity.Location.PropertyPath ?? string.Empty);

    private static bool ExcludeGeneratedUnid(XAttribute attribute) =>
        attribute.Name != PtOpenXml.Unid;

    private static bool IncludeRevisionAttribute(XAttribute attribute) =>
        ExcludeGeneratedUnid(attribute)
        && !(attribute.Name.LocalName == "id"
            && IsWordNamespace(attribute.Name.NamespaceName));

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
        var normalized = part.Xml is null
            ? part.RawBytesDigest!
            : XmlSemanticNormalizer.Digest(
                part.Xml,
                partUri,
                ignoreFormattingWhitespace: true,
                includeAttribute: ExcludeGeneratedUnid,
                attributeValueNormalizer: RelationshipAttributeNormalizer(
                    partUri, relationships));
        string fingerprint = normalized.Value;
        var contentType = ContentTypeFor(part);
        var digest = SemanticValue.Digest(
            normalized.Algorithm,
            fingerprint,
            part.Xml is null ? "raw-part-bytes" : "xml-expanded-names-comments-pi-v1");
        var value = part.Xml is null
            ? ValueObj(
                ("contentType", SemanticValue.String(contentType)),
                ("size", SemanticValue.Integer(part.Size)),
                ("normalizedDigest", digest))
            : ValueObj(
                ("contentType", SemanticValue.String(contentType)),
                ("normalizedDigest", digest));
        var identity = ValueObj(
            ("contentType", SemanticValue.String(contentType)),
            ("normalizedDigest", digest));
        return new Entity(
            "registry:" + partUri,
            family,
            new ChangeLocation
            {
                EntryUri = partUri,
                PropertyPath = path,
            },
            null,
            null,
            value,
            ValueFingerprint(identity));
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

    private static string ContentTypeFor(Part part)
    {
        if (part.ContentType is not null)
            return part.ContentType;
        var extension = Path.GetExtension(part.Name).ToLowerInvariant();
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

    private sealed record Part(
        string Name,
        XDocument? Xml,
        string? ContentType,
        long Size,
        VerificationDigest? RawBytesDigest);

    private sealed record RelationshipInfo(
        string Owner,
        string Id,
        string Type,
        string Target,
        string Mode);

    private sealed record RelationshipReadResult(
        IReadOnlyList<Entity> Inventory,
        IReadOnlyDictionary<(string Owner, string Id), RelationshipInfo> ByOwnerAndId);

    private record Entity(
        string Key,
        SemanticChangeFamily Family,
        ChangeLocation Location,
        string? Anchor,
        string? Scope,
        SemanticValue Value,
        string Fingerprint,
        string? LocationKey = null);

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
