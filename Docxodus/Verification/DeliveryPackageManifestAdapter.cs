// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Globalization;
using System.Text;

namespace Docxodus.Verification;

/// <summary>
/// The single #456 integration point for delivery receipts. Package parsing remains owned by
/// PackageManifestGenerator; this adapter only projects the public manifest into immutable receipt
/// identities and compares already-materialized records. Corrected #493 availability semantics for
/// optional content/semantic digests remain isolated here rather than being reinterpreted by receipts.
/// </summary>
internal static class DeliveryPackageManifestAdapter
{
    public static bool IsSupportedSchema(string schema) =>
        string.Equals(schema, PackageManifest.SchemaId, StringComparison.Ordinal);

    public static DeliveryDocumentIdentity CreateIdentity(
        PackageManifest manifest,
        long documentVersion)
    {
        ArgumentNullException.ThrowIfNull(manifest);
        if (documentVersion < 0)
            throw new ArgumentOutOfRangeException(nameof(documentVersion));
        if (!IsSupportedSchema(manifest.Schema) || manifest.SchemaVersion != 1)
        {
            throw new DeliveryReceiptValidationException(
                "unsupported_package_manifest",
                $"Expected {PackageManifest.SchemaId} version 1.");
        }
        if (!manifest.IsValid)
        {
            throw new DeliveryReceiptValidationException(
                "invalid_package_manifest",
                "Delivery document identities require a manifest with no error findings.");
        }
        DeliveryReceiptValidation.ValidateDigest(manifest.RawPackageBytesDigest, "raw package digest");
        DeliveryReceiptValidation.ValidateOptionalDigest(
            manifest.OrderedOpcContentDigest, "ordered OPC content digest");
        DeliveryReceiptValidation.ValidateOptionalDigest(
            manifest.NormalizedSemanticDigest, "normalized semantic digest");
        var packageKind = DeliveryReceiptValidation.RequireNonBlank(
            manifest.PackageKind, "package kind", 256);
        return new DeliveryDocumentIdentity
        {
            DocumentVersion = documentVersion,
            PackageKind = packageKind,
            PackageManifestSchema = manifest.Schema,
            RawPackageBytesDigest = DeliveryReceiptValidation.CloneDigest(
                manifest.RawPackageBytesDigest),
            OrderedOpcContentDigest = DeliveryReceiptValidation.CloneOptionalDigest(
                manifest.OrderedOpcContentDigest),
            NormalizedSemanticDigest = DeliveryReceiptValidation.CloneOptionalDigest(
                manifest.NormalizedSemanticDigest),
        };
    }

    public static IReadOnlyList<DeliveryPackageChangeObservation> Compare(
        PackageManifest before,
        PackageManifest after)
    {
        ArgumentNullException.ThrowIfNull(before);
        ArgumentNullException.ThrowIfNull(after);
        var changes = new List<DeliveryPackageChangeObservation>();
        AddEntryChanges(changes, before.Entries, after.Entries);
        AddRelationshipChanges(changes, before.Relationships, after.Relationships);
        return changes
            .OrderBy(change => change.Kind)
            .ThenBy(change => change.Location.EntryUri, StringComparer.Ordinal)
            .ThenBy(change => change.Location.OwnerUri, StringComparer.Ordinal)
            .ThenBy(change => change.Location.RelationshipId, StringComparer.Ordinal)
            .ThenBy(change => change.Location.PropertyPath, StringComparer.Ordinal)
            .ToArray();
    }

    private static void AddEntryChanges(
        ICollection<DeliveryPackageChangeObservation> changes,
        IReadOnlyList<PackageManifestEntry> before,
        IReadOnlyList<PackageManifestEntry> after)
    {
        var left = before.ToDictionary(
            entry => (entry.Uri, entry.Occurrence), EntryValue);
        var right = after.ToDictionary(
            entry => (entry.Uri, entry.Occurrence), EntryValue);
        foreach (var key in left.Keys.Union(right.Keys)
                     .OrderBy(key => key.Uri, StringComparer.Ordinal)
                     .ThenBy(key => key.Occurrence))
        {
            left.TryGetValue(key, out var oldValue);
            right.TryGetValue(key, out var newValue);
            if (string.Equals(oldValue, newValue, StringComparison.Ordinal))
                continue;
            changes.Add(new DeliveryPackageChangeObservation(
                oldValue is null ? DeliveryPackageChangeKind.PartAdded
                    : newValue is null ? DeliveryPackageChangeKind.PartRemoved
                    : DeliveryPackageChangeKind.PartModified,
                new ChangeLocation
                {
                    EntryUri = key.Uri,
                    PropertyPath = $"occurrence:{key.Occurrence.ToString(CultureInfo.InvariantCulture)}",
                },
                oldValue,
                newValue));
        }
    }

    private static void AddRelationshipChanges(
        ICollection<DeliveryPackageChangeObservation> changes,
        IReadOnlyList<PackageRelationship> before,
        IReadOnlyList<PackageRelationship> after)
    {
        var left = IndexedRelationships(before);
        var right = IndexedRelationships(after);
        foreach (var key in left.Keys.Union(right.Keys)
                     .OrderBy(key => key.OwnerUri, StringComparer.Ordinal)
                     .ThenBy(key => key.Id, StringComparer.Ordinal)
                     .ThenBy(key => key.Occurrence))
        {
            left.TryGetValue(key, out var oldRelationship);
            right.TryGetValue(key, out var newRelationship);
            var oldValue = oldRelationship is null ? null : RelationshipValue(oldRelationship);
            var newValue = newRelationship is null ? null : RelationshipValue(newRelationship);
            if (string.Equals(oldValue, newValue, StringComparison.Ordinal))
                continue;
            changes.Add(new DeliveryPackageChangeObservation(
                oldRelationship is null ? DeliveryPackageChangeKind.RelationshipAdded
                    : newRelationship is null ? DeliveryPackageChangeKind.RelationshipRemoved
                    : DeliveryPackageChangeKind.RelationshipModified,
                new ChangeLocation
                {
                    OwnerUri = key.OwnerUri,
                    RelationshipId = key.Id,
                    TargetUri = newRelationship?.ResolvedTargetUri ?? newRelationship?.Target
                        ?? oldRelationship?.ResolvedTargetUri ?? oldRelationship?.Target,
                    PropertyPath = $"occurrence:{key.Occurrence.ToString(CultureInfo.InvariantCulture)}",
                },
                oldValue,
                newValue));
        }
    }

    private static string EntryValue(PackageManifestEntry entry) =>
        Encoding.UTF8.GetString(DeliveryReceiptCanonicalJson.SerializeCanonical(new
        {
            contentType = entry.ContentType,
            contentTypeSource = entry.ContentTypeSource,
            isEncrypted = entry.IsEncrypted,
            isXml = entry.IsXml,
            normalizedXmlDigest = entry.NormalizedXmlDigest,
            rawBytesDigest = entry.RawBytesDigest,
            size = entry.Size,
        }));

    private static string RelationshipValue(PackageRelationship relationship) =>
        Encoding.UTF8.GetString(DeliveryReceiptCanonicalJson.SerializeCanonical(new
        {
            isTargetPresent = relationship.IsTargetPresent,
            resolvedTargetUri = relationship.ResolvedTargetUri,
            target = relationship.Target,
            targetMode = relationship.TargetMode,
            type = relationship.Type,
        }));

    private static Dictionary<(string OwnerUri, string Id, int Occurrence), PackageRelationship>
        IndexedRelationships(IReadOnlyList<PackageRelationship> relationships)
    {
        var result = new Dictionary<(string, string, int), PackageRelationship>();
        foreach (var group in relationships
                     .GroupBy(relationship => (relationship.OwnerUri, relationship.Id)))
        {
            var occurrence = 0;
            foreach (var relationship in group.OrderBy(RelationshipValue, StringComparer.Ordinal))
                result.Add((group.Key.OwnerUri, group.Key.Id, occurrence++), relationship);
        }
        return result;
    }
}

internal sealed record DeliveryPackageChangeObservation(
    DeliveryPackageChangeKind Kind,
    ChangeLocation Location,
    string? Before,
    string? After);
