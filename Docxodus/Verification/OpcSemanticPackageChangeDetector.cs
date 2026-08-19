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
        CompareEntities(left.ContentTypes, right.ContentTypes, result);
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

        var relationshipData = ReadRelationships(inspection.Manifest.Relationships, parts);

        // Relationship bindings, bookmarks, and revisions all use the nearest deterministic Word
        // anchor as their stable location axis. Assign those anchors before reading any of them so
        // inserting an unrelated preceding block cannot turn unchanged package facts into moves.
        foreach (var part in parts.Values
            .Where(item => item.Xml?.Root is not null && IsWordStoryPart(item)))
            UnidHelper.AssignToAllElementsDeterministic(part.Xml!.Root!);

        var relationshipBindings = ReadRelationshipBindings(parts, relationshipData.ByOwnerAndId);
        var bookmarks = new List<Entity>();
        var revisions = new List<Entity>();
        var annotations = new List<Entity>();
        var storyResiduals = new List<Entity>();
        foreach (var part in parts.Values.OrderBy(item => item.Name, StringComparer.Ordinal))
        {
            if (part.Xml?.Root is null) continue;
            var partUri = PartUri(part.Name);
            if (IsWordStoryPart(part))
            {
                bookmarks.AddRange(ReadBookmarks(part, partUri));
                revisions.AddRange(ReadRevisions(
                    part, partUri, relationshipData.ByOwnerAndId));
                var residual = StoryEnvelopeEntity(
                    part, partUri, relationshipData.ByOwnerAndId);
                if (residual is not null) storyResiduals.Add(residual);
                var extensionResidual = StoryExtensionEntity(
                    part, partUri, relationshipData.ByOwnerAndId);
                if (extensionResidual is not null) storyResiduals.Add(extensionResidual);
            }
            if (part.Xml.Root.Name.NamespaceName == AnnotationNamespace
                && part.Xml.Root.Name.LocalName == "annotations")
                annotations.AddRange(ReadAnnotations(
                    part, partUri, relationshipData.ByOwnerAndId));
        }

        var media = parts.Values
            .Where(IsMediaPart)
            .OrderBy(part => part.Name, StringComparer.Ordinal)
            .Select(part =>
            {
                var value = ValueObj(
                    ("contentType", SemanticValue.String(ContentTypeFor(part))),
                    ("size", SemanticValue.IntegerFromDocument(part.Size)),
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

        // Content-type declarations are package semantics even for otherwise recognized story
        // parts. Comparing both the resolved per-entry role and the canonical declaration set
        // covers role/MIME changes plus unused declarations without depending on XML order.
        var contentTypes = inspection.Manifest.Entries
            .Where(entry =>
            {
                if (entry.Uri.EndsWith("/", StringComparison.Ordinal)) return false;
                if (!parts.TryGetValue(entry.Uri.TrimStart('/'), out var part)) return false;
                return IsWordStoryPart(part)
                    || part.Xml?.Root?.Name is
                        { NamespaceName: AnnotationNamespace, LocalName: "annotations" };
            })
            .OrderBy(entry => entry.Uri, StringComparer.Ordinal)
            .Select(entry =>
            {
                var part = parts[entry.Uri.TrimStart('/')];
                var contentType = CanonicalContentType(entry.ContentType);
                var value = ValueObj(("contentType", SemanticValue.String(contentType)));
                var key = "content-type:resolved:" + entry.Uri;
                return new Entity(
                    key,
                    SemanticChangeFamily.OpaquePackagePart,
                    new ChangeLocation
                    {
                        EntryUri = entry.Uri,
                        PropertyPath = "package.content_type",
                    },
                    null,
                    ScopeForPart(part),
                    value,
                    ValueFingerprint(value),
                    entry.Uri,
                    key);
            })
            .Concat(inspection.Manifest.ContentTypes
                .Where(declaration => !ContentTypeDeclarationIsUsed(
                    declaration, inspection.Manifest.Entries))
                .OrderBy(declaration => declaration.Kind, StringComparer.Ordinal)
                .ThenBy(declaration => declaration.Key, StringComparer.OrdinalIgnoreCase)
                .ThenBy(declaration => declaration.ContentType, StringComparer.OrdinalIgnoreCase)
                .Select(declaration =>
                {
                    var canonicalKey = declaration.Kind == "default"
                        ? declaration.Key.ToLowerInvariant()
                        : declaration.Key;
                    var groupKey = string.Join(":",
                        "content-type", "declaration", declaration.Kind, canonicalKey);
                    var value = ValueObj(
                        ("kind", SemanticValue.String(declaration.Kind)),
                        ("key", SemanticValue.String(canonicalKey)),
                        ("contentType", SemanticValue.String(
                            CanonicalContentType(declaration.ContentType))));
                    return new Entity(
                        groupKey,
                        SemanticChangeFamily.OpaquePackagePart,
                        new ChangeLocation
                        {
                            EntryUri = "/[Content_Types].xml",
                            PropertyPath = $"package.content_type.declaration[{declaration.Kind}:{canonicalKey}]",
                        },
                        null,
                        null,
                        value,
                        ValueFingerprint(value),
                        groupKey,
                        groupKey);
                }))
            .ToArray();

        var registryParts = parts.Values
            .Where(IsRegistryPart)
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
                        ("size", SemanticValue.IntegerFromDocument(part.Size)),
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
            .Concat(storyResiduals)
            .ToArray();

        return new PackageSnapshot(
            contentTypes,
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
        IReadOnlyList<PackageRelationship> relationships,
        IReadOnlyDictionary<string, Part> parts)
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
            parts.TryGetValue(relationship.OwnerUri.TrimStart('/'), out var ownerPart);
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
                ownerPart is null
                    ? ScopeForPart(relationship.OwnerUri)
                    : ScopeForPart(ownerPart),
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
                var anchorElement = NearestAnchorElement(element);
                var elementPath = anchorElement is null
                    ? ElementPath(root, element)
                    : RelativeElementPath(anchorElement, element);
                var attributeName = ExpandedName(attribute.Name);
                var anchor = AnchorFor(anchorElement, part);
                var locationKey = string.Join("\u001f",
                    anchor ?? ScopeForPart(part) ?? owner,
                    elementPath,
                    attributeName);
                var key = $"relationship-binding:{owner}:{locationKey}";
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
                    ScopeForPart(part),
                    value,
                    ValueFingerprint(value),
                    locationKey,
                    "relationship.binding:" + attributeName));
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
            .Where(element => WordAttr(element, "id") is not null)
            .GroupBy(element => WordAttr(element, "id")!, StringComparer.Ordinal)
            .ToDictionary(
                group => group.Key,
                group => new Queue<XElement>(group),
                StringComparer.Ordinal);
        var starts = part.Xml!.Descendants()
            .Where(element => element.Name.LocalName == "bookmarkStart"
                && IsWordNamespace(element.Name.NamespaceName))
            .ToArray();
        var grouped = starts.GroupBy(element => WordAttr(element, "name") ?? string.Empty)
            .OrderBy(group => group.Key, StringComparer.Ordinal);
        foreach (var group in grouped)
        {
            int ordinal = 0;
            foreach (var bookmark in group)
            {
                var anchorElement = NearestAnchorElement(bookmark);
                var anchor = AnchorFor(anchorElement, part);
                var name = WordAttr(bookmark, "name") ?? string.Empty;
                var nativeId = WordAttr(bookmark, "id");
                XElement? end = null;
                if (nativeId is not null && endsById.TryGetValue(nativeId, out var candidates)
                    && candidates.Count > 0)
                    end = candidates.Dequeue();
                var endAnchorElement = end is null ? null : NearestAnchorElement(end);
                var endAnchor = AnchorFor(endAnchorElement, part);
                var startPath = anchorElement is null
                    ? ElementPath(root, bookmark)
                    : RelativeElementPath(anchorElement, bookmark);
                var endPath = end is null ? null
                    : endAnchorElement is null ? ElementPath(root, end)
                    : RelativeElementPath(endAnchorElement, end);
                var value = ValueObj(
                    ("name", SemanticValue.String(name)),
                    ("columnFirst", SemanticValue.IntegerFromDocument(ParseLong(WordAttr(bookmark, "colFirst")))),
                    ("columnLast", SemanticValue.IntegerFromDocument(ParseLong(WordAttr(bookmark, "colLast")))),
                    ("startAnchor", SemanticValue.String(anchor)),
                    ("endAnchor", SemanticValue.String(endAnchor)),
                    ("startPath", SemanticValue.String(startPath)),
                    ("endPath", SemanticValue.String(endPath)));
                var fingerprint = ValueFingerprint(ValueObj(
                    ("name", SemanticValue.String(name)),
                    ("columnFirst", SemanticValue.IntegerFromDocument(ParseLong(WordAttr(bookmark, "colFirst")))),
                    ("columnLast", SemanticValue.IntegerFromDocument(ParseLong(WordAttr(bookmark, "colLast"))))));
                yield return new Entity(
                    $"bookmark:{partUri}:{name}:{ordinal++}",
                    SemanticChangeFamily.Bookmark,
                    new ChangeLocation
                    {
                        EntryUri = partUri,
                        PropertyPath = "bookmark",
                    },
                    anchor,
                    ScopeForPart(part),
                    value,
                    fingerprint,
                    string.Join("\u001f", anchor, startPath, endAnchor, endPath),
                    "bookmark");
            }
        }
    }

    private static IEnumerable<Entity> ReadRevisions(
        Part part,
        string partUri,
        IReadOnlyDictionary<(string Owner, string Id), RelationshipInfo> relationships)
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
            var nativeId = WordAttr(revision, "id");
            var identity = nativeId is null ? kind : kind + ":" + nativeId;
            ordinals.TryGetValue(identity, out var ordinal);
            ordinals[identity] = ordinal + 1;
            var anchorElement = NearestAnchorElement(revision);
            var anchor = AnchorFor(anchorElement, part);
            var structuralPath = anchorElement is null
                ? ElementPath(root, revision)
                : RelativeElementPath(anchorElement, revision);
            var normalizedRevision = XmlSemanticNormalizer.Digest(
                revision,
                partUri,
                ignoreFormattingWhitespace: true,
                includeAttribute: IncludeRevisionAttribute,
                attributeValueNormalizer: RelationshipAttributeNormalizer(
                    partUri, relationships));
            var value = ValueObj(
                ("kind", SemanticValue.String(kind)),
                ("author", SemanticValue.String(WordAttr(revision, "author"))),
                ("date", SemanticValue.String(WordAttr(revision, "date"))),
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
                ScopeForPart(part),
                value,
                ValueFingerprint(value),
                string.Join("\u001f", anchor, structuralPath),
                "revision:" + kind);
        }
    }

    private static IEnumerable<Entity> ReadAnnotations(
        Part part,
        string partUri,
        IReadOnlyDictionary<(string Owner, string Id), RelationshipInfo> relationships)
    {
        foreach (var annotation in part.Xml!.Root!.Elements()
            .Where(element => element.Name.NamespaceName == AnnotationNamespace
                && element.Name.LocalName == "annotation")
            .OrderBy(element => UnqualifiedAttr(element, "id"), StringComparer.Ordinal))
        {
            var id = UnqualifiedAttr(annotation, "id") ?? string.Empty;
            var bookmarkName = annotation.Descendants()
                .FirstOrDefault(element => element.Name.NamespaceName == AnnotationNamespace
                    && element.Name.LocalName == "range")?
                .Attribute("bookmarkName")?.Value;
            var normalized = XmlSemanticNormalizer.Digest(
                annotation,
                partUri,
                ignoreFormattingWhitespace: false,
                includeAttribute: ExcludeGeneratedUnid,
                attributeValueNormalizer: RelationshipAttributeNormalizer(
                    partUri, relationships));
            var value = ValueObj(
                ("id", SemanticValue.String(id)),
                ("labelId", SemanticValue.String(UnqualifiedAttr(annotation, "labelId"))),
                ("label", SemanticValue.String(UnqualifiedAttr(annotation, "label"))),
                ("color", SemanticValue.String(UnqualifiedAttr(annotation, "color"))),
                ("author", SemanticValue.String(UnqualifiedAttr(annotation, "author"))),
                ("created", SemanticValue.String(UnqualifiedAttr(annotation, "created"))),
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
                $"{partUri}\u001f{id}",
                "annotation");
        }

        var root = part.Xml.Root!;
        var envelope = ResidualEnvelope(root, node =>
            node is XElement element
            && element.Name.NamespaceName == AnnotationNamespace
            && element.Name.LocalName == "annotation");
        var envelopeDigest = XmlSemanticNormalizer.Digest(
            ResidualDocument(part.Xml, root, envelope),
            partUri,
            ignoreFormattingWhitespace: true,
            includeAttribute: ExcludeGeneratedUnid,
            attributeValueNormalizer: RelationshipAttributeNormalizer(
                partUri, relationships));
        var envelopeValue = ValueObj(("normalizedDigest", SemanticValue.Digest(
            envelopeDigest.Algorithm,
            envelopeDigest.Value,
            "docxodus-annotation-envelope-v1")));
        yield return new Entity(
            "annotation-envelope:" + partUri,
            SemanticChangeFamily.Annotation,
            new ChangeLocation
            {
                EntryUri = partUri,
                PropertyPath = "annotation.registry.package",
            },
            null,
            null,
            envelopeValue,
            envelopeDigest.Value,
            partUri,
            "annotation.registry.package");
    }

    private static Entity? StoryEnvelopeEntity(
        Part part,
        string partUri,
        IReadOnlyDictionary<(string Owner, string Id), RelationshipInfo> relationships)
    {
        var root = part.Xml!.Root!;
        bool IsModeledRootChild(XNode node)
        {
            if (node is not XElement element) return false;
            if (!IsWordNamespace(element.Name.NamespaceName)) return false;
            return root.Name.LocalName switch
            {
                "document" => element.Name.LocalName == "body",
                "footnotes" => element.Name.LocalName == "footnote",
                "endnotes" => element.Name.LocalName == "endnote",
                "comments" => element.Name.LocalName == "comment",
                // Header/footer direct elements are all fed through the IR block reader, whose
                // opaque block type retains extension elements.
                "hdr" or "ftr" => true,
                _ => false,
            };
        }

        bool hasSemanticEnvelope = root.Attributes().Any(attribute =>
                !attribute.IsNamespaceDeclaration && ExcludeGeneratedUnid(attribute))
            || root.Nodes().Any(node => IsMeaningfulResidualNode(
                node, IsModeledRootChild))
            || part.Xml.Nodes().Any(node => node != root
                && IsMeaningfulResidualNode(node, _ => false));
        if (!hasSemanticEnvelope) return null;

        var envelope = ResidualEnvelope(root, IsModeledRootChild);
        var normalized = XmlSemanticNormalizer.Digest(
            ResidualDocument(part.Xml, root, envelope),
            partUri,
            ignoreFormattingWhitespace: true,
            includeAttribute: ExcludeGeneratedUnid,
            attributeValueNormalizer: RelationshipAttributeNormalizer(
                partUri, relationships));
        var value = ValueObj(("normalizedDigest", SemanticValue.Digest(
            normalized.Algorithm,
            normalized.Value,
            "word-story-envelope-v1")));
        return new Entity(
            "story-envelope:" + partUri,
            SemanticChangeFamily.OpaquePackagePart,
            new ChangeLocation
            {
                EntryUri = partUri,
                PropertyPath = "story.envelope.package",
            },
            null,
            ScopeForPart(part),
            value,
            normalized.Value,
            partUri,
            "story.envelope.package");
    }

    private static Entity? StoryExtensionEntity(
        Part part,
        string partUri,
        IReadOnlyDictionary<(string Owner, string Id), RelationshipInfo> relationships)
    {
        var root = part.Xml!.Root!;
        var residualNamespace = XNamespace.Get("urn:docxodus:semantic-story-residual:v1");
        var records = new List<XElement>();

        // The IR is intentionally total for unknown elements, but attributes and XML node kinds
        // attached to otherwise modeled Word paragraphs/runs are not necessarily represented by a
        // typed IR field. Keep only those residual facts here: hashing modeled text/properties again
        // would create a noisy opaque change alongside every ordinary edit.
        foreach (var element in root.Descendants())
        {
            var extensionAttributes = element.Attributes()
                .Where(IsStoryExtensionAttribute)
                .ToArray();
            var semanticNodes = element.Nodes()
                .Select((node, position) => (Node: node, Position: position))
                .Where(item => item.Node is XComment or XProcessingInstruction)
                .ToArray();
            if (extensionAttributes.Length == 0 && semanticNodes.Length == 0) continue;

            var anchorElement = NearestAnchorElement(element);
            var anchor = AnchorFor(anchorElement, part);
            var path = anchorElement is null
                ? ElementPath(root, element)
                : RelativeElementPath(anchorElement, element);

            foreach (var attribute in extensionAttributes)
            {
                var source = new XElement(
                    element.Name,
                    InScopeNamespaceDeclarations(element),
                    new XAttribute(attribute));
                records.Add(new XElement(
                    residualNamespace + "attribute",
                    new XAttribute("anchor", anchor ?? ScopeForPart(part) ?? partUri),
                    new XAttribute("path", path),
                    new XAttribute("name", ExpandedName(attribute.Name)),
                    source));
            }

            foreach (var (node, position) in semanticNodes)
            {
                records.Add(new XElement(
                    residualNamespace + "node",
                    new XAttribute("anchor", anchor ?? ScopeForPart(part) ?? partUri),
                    new XAttribute("path", path),
                    new XAttribute("ordinal", position),
                    CloneNode(node)));
            }
        }

        if (records.Count == 0) return null;
        var orderedRecords = records
            .OrderBy(record => record.Name.LocalName, StringComparer.Ordinal)
            .ThenBy(record => (string?)record.Attribute("anchor"), StringComparer.Ordinal)
            .ThenBy(record => (string?)record.Attribute("path"), StringComparer.Ordinal)
            .ThenBy(record => (string?)record.Attribute("name"), StringComparer.Ordinal)
            .ThenBy(record => (int?)record.Attribute("ordinal"))
            .ToArray();
        var normalized = XmlSemanticNormalizer.Digest(
            new XDocument(new XElement(residualNamespace + "story", orderedRecords)),
            partUri,
            ignoreFormattingWhitespace: false,
            includeAttribute: ExcludeGeneratedUnid,
            attributeValueNormalizer: RelationshipAttributeNormalizer(
                partUri, relationships));
        var value = ValueObj(("normalizedDigest", SemanticValue.Digest(
            normalized.Algorithm,
            normalized.Value,
            "word-story-extension-residual-v1")));
        return new Entity(
            "story-extensions:" + partUri,
            SemanticChangeFamily.OpaquePackagePart,
            new ChangeLocation
            {
                EntryUri = partUri,
                PropertyPath = "story.extensions.package",
            },
            null,
            ScopeForPart(part),
            value,
            normalized.Value,
            partUri,
            "story.extensions.package");
    }

    private static bool IsStoryExtensionAttribute(XAttribute attribute)
    {
        if (attribute.IsNamespaceDeclaration || !ExcludeGeneratedUnid(attribute)) return false;
        var namespaceName = attribute.Name.NamespaceName;
        return namespaceName.Length > 0
            && !IsWordNamespace(namespaceName)
            && namespaceName != OfficeRelationshipNamespace
            && namespaceName != StrictOfficeRelationshipNamespace
            && namespaceName != XNamespace.Xml.NamespaceName;
    }

    private static IEnumerable<XAttribute> InScopeNamespaceDeclarations(XElement element)
    {
        var declarations = new Dictionary<XName, string>();
        foreach (var ancestor in element.AncestorsAndSelf().Reverse())
        {
            foreach (var attribute in ancestor.Attributes().Where(item => item.IsNamespaceDeclaration))
                declarations[attribute.Name] = attribute.Value;
        }
        return declarations
            .OrderBy(item => item.Key.NamespaceName, StringComparer.Ordinal)
            .ThenBy(item => item.Key.LocalName, StringComparer.Ordinal)
            .Select(item => new XAttribute(item.Key, item.Value));
    }

    private static XElement ResidualEnvelope(
        XElement root,
        Func<XNode, bool> excludeNode) => new(
            root.Name,
            root.Attributes().Select(attribute => new XAttribute(attribute)),
            root.Nodes().Where(node => !excludeNode(node)).Select(CloneNode));

    private static XDocument ResidualDocument(
        XDocument source,
        XElement sourceRoot,
        XElement residualRoot) => new(
            source.Nodes().Select(node => node == sourceRoot ? residualRoot : CloneNode(node)));

    private static bool IsMeaningfulResidualNode(
        XNode node,
        Func<XNode, bool> excludeNode) => !excludeNode(node)
            && (node is not XText text || !string.IsNullOrWhiteSpace(text.Value));

    private static XNode CloneNode(XNode node) => node switch
    {
        XElement element => new XElement(element),
        XComment comment => new XComment(comment.Value),
        XProcessingInstruction instruction =>
            new XProcessingInstruction(instruction.Target, instruction.Data),
        XCData cdata => new XCData(cdata.Value),
        XText text => new XText(text.Value),
        _ => throw new InvalidDataException(
            $"Unsupported XML node type '{node.NodeType}' in semantic residual."),
    };

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
        entity.GroupKey ?? entity.Location.PropertyPath ?? string.Empty);

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
        var family = IsNumberingRegistryPart(part)
            ? SemanticChangeFamily.Numbering
            : SemanticChangeFamily.Style;
        var path = IsNumberingRegistryPart(part)
            ? "numbering.registry.package"
            : IsThemeRegistryPart(part)
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
                ("size", SemanticValue.IntegerFromDocument(part.Size)),
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
        if (IsMediaPart(part) || IsWordStoryPart(part)) return false;
        if (IsRegistryPart(part)) return false;
        if (part.Xml?.Root?.Name is { NamespaceName: AnnotationNamespace, LocalName: "annotations" })
            return false;
        return true;
    }

    private static bool IsRegistryPart(Part part)
    {
        var name = part.Name;
        var contentType = ContentTypeEssence(part);
        return name is "word/styles.xml" or "word/numbering.xml"
            || name.StartsWith("word/theme/", StringComparison.Ordinal)
            || contentType.Contains("wordprocessingml.styles+xml", StringComparison.Ordinal)
            || contentType.Contains("wordprocessingml.numbering+xml", StringComparison.Ordinal)
            || contentType.Contains("officedocument.theme+xml", StringComparison.Ordinal);
    }

    private static bool IsNumberingRegistryPart(Part part) =>
        part.Name == "word/numbering.xml"
        || ContentTypeEssence(part).Contains(
            "wordprocessingml.numbering+xml", StringComparison.Ordinal);

    private static bool IsThemeRegistryPart(Part part) =>
        part.Name.StartsWith("word/theme/", StringComparison.Ordinal)
        || ContentTypeEssence(part).Contains(
            "officedocument.theme+xml", StringComparison.Ordinal);

    private static bool IsWordStoryPart(Part part)
    {
        var name = part.Name;
        if (name == "word/document.xml"
            || name == "word/footnotes.xml"
            || name == "word/endnotes.xml"
            || name == "word/comments.xml"
            || (name.StartsWith("word/header", StringComparison.Ordinal)
                && name.EndsWith(".xml", StringComparison.Ordinal))
            || (name.StartsWith("word/footer", StringComparison.Ordinal)
                && name.EndsWith(".xml", StringComparison.Ordinal)))
            return true;

        var contentType = ContentTypeEssence(part);
        return contentType.Contains("wordprocessingml.document.main+xml", StringComparison.Ordinal)
            || contentType.Contains("wordprocessingml.header+xml", StringComparison.Ordinal)
            || contentType.Contains("wordprocessingml.footer+xml", StringComparison.Ordinal)
            || contentType.Contains("wordprocessingml.footnotes+xml", StringComparison.Ordinal)
            || contentType.Contains("wordprocessingml.endnotes+xml", StringComparison.Ordinal)
            || contentType.Contains("wordprocessingml.comments+xml", StringComparison.Ordinal)
            || contentType.Contains("ms-word.document.macroenabled.main+xml", StringComparison.Ordinal)
            || contentType.Contains("ms-word.template.macroenabledtemplate.main+xml", StringComparison.Ordinal);
    }

    private static bool IsMediaPart(Part part)
    {
        var contentType = ContentTypeEssence(part);
        if (contentType.StartsWith("image/", StringComparison.Ordinal)
            || contentType.StartsWith("audio/", StringComparison.Ordinal)
            || contentType.StartsWith("video/", StringComparison.Ordinal))
            return true;
        if (IsXmlContentType(contentType)) return false;
        var name = part.Name;
        return name.StartsWith("media/", StringComparison.Ordinal)
            || name.Contains("/media/", StringComparison.Ordinal)
            || name.StartsWith("word/embeddings/", StringComparison.Ordinal);
    }

    private static string ContentTypeEssence(Part part)
    {
        var value = part.ContentType ?? string.Empty;
        var semicolon = value.IndexOf(';');
        return (semicolon < 0 ? value : value[..semicolon]).ToLowerInvariant();
    }

    private static bool IsXmlContentType(string contentType) =>
        contentType is "application/xml" or "text/xml"
        || contentType.EndsWith("+xml", StringComparison.Ordinal);

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

    private static string RelativeElementPath(XElement anchor, XElement element)
    {
        if (anchor == element) return ".";
        var segments = new Stack<string>();
        XElement? current = element;
        while (current is not null && current != anchor)
        {
            int ordinal = current.ElementsBeforeSelf()
                .Count(sibling => sibling.Name == current.Name) + 1;
            segments.Push($"{ExpandedName(current.Name)}[{ordinal}]");
            current = current.Parent;
        }
        return current == anchor
            ? "./" + string.Join("/", segments)
            : ElementPath(anchor.Document?.Root ?? anchor, element);
    }

    private static string ExpandedName(XName name) =>
        $"{{{name.NamespaceName}}}{name.LocalName}";

    private static XElement? NearestAnchorElement(XElement element)
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
            return candidate;
        }
        return null;
    }

    private static string? AnchorFor(XElement? element, Part part)
    {
        if (element is null) return null;
        string? kind = element.Name.LocalName switch
        {
            "p" => "p",
            "tbl" => "tbl",
            "tr" => "tr",
            "tc" => "tc",
            "sdt" => "sdt",
            "sectPr" => "sec",
            _ => null,
        };
        var unid = (string?)element.Attribute(PtOpenXml.Unid);
        return kind is null || string.IsNullOrWhiteSpace(unid)
            ? null
            : $"{kind}:{ScopeForPart(part)}:{unid}";
    }

    private static string? ScopeForPart(Part part)
    {
        var pathScope = ScopeForPart(PartUri(part.Name));
        if (pathScope is not null) return pathScope;
        var contentType = ContentTypeEssence(part);
        if (contentType.Contains("wordprocessingml.document.main+xml", StringComparison.Ordinal)
            || contentType.Contains("ms-word.document.macroenabled.main+xml", StringComparison.Ordinal)
            || contentType.Contains("ms-word.template.macroenabledtemplate.main+xml", StringComparison.Ordinal))
            return "body";
        if (contentType.Contains("wordprocessingml.footnotes+xml", StringComparison.Ordinal)) return "fn";
        if (contentType.Contains("wordprocessingml.endnotes+xml", StringComparison.Ordinal)) return "en";
        if (contentType.Contains("wordprocessingml.comments+xml", StringComparison.Ordinal)) return "cmt";
        if (contentType.Contains("wordprocessingml.header+xml", StringComparison.Ordinal)) return "hdr";
        if (contentType.Contains("wordprocessingml.footer+xml", StringComparison.Ordinal)) return "ftr";
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

    private static string? WordAttr(XElement element, string localName) =>
        IsWordNamespace(element.Name.NamespaceName)
            ? (string?)element.Attribute(XName.Get(localName, element.Name.NamespaceName))
            : null;

    private static string? UnqualifiedAttr(XElement element, string localName) =>
        (string?)element.Attribute(XName.Get(localName));

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

    private static string? CanonicalContentType(string? value)
    {
        if (value is null) return null;
        // MIME type/subtype casing is insensitive, but parameter values are not generally so.
        // Preserve the validated parameter suffix verbatim rather than hiding a meaningful value
        // change (for example, a case-sensitive profile or boundary token).
        int parameter = value.IndexOf(';');
        return parameter < 0
            ? value.ToLowerInvariant()
            : value[..parameter].ToLowerInvariant() + value[parameter..];
    }

    private static bool ContentTypeDeclarationIsUsed(
        PackageContentTypeDeclaration declaration,
        IReadOnlyList<PackageManifestEntry> entries)
    {
        if (declaration.Kind == "override")
            return entries.Any(entry => string.Equals(
                entry.Uri, declaration.Key, StringComparison.OrdinalIgnoreCase));
        return entries.Any(entry => string.Equals(
            Path.GetExtension(entry.Uri).TrimStart('.'),
            declaration.Key,
            StringComparison.OrdinalIgnoreCase));
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
        string? LocationKey = null,
        string? GroupKey = null);

    private sealed record PackageSnapshot(
        IReadOnlyList<Entity> ContentTypes,
        IReadOnlyList<Entity> Relationships,
        IReadOnlyList<Entity> RelationshipBindings,
        IReadOnlyList<Entity> Media,
        IReadOnlyList<Entity> Bookmarks,
        IReadOnlyList<Entity> Revisions,
        IReadOnlyList<Entity> Annotations,
        IReadOnlyList<Entity> RegistryParts,
        IReadOnlyList<Entity> OpaqueParts);
}
