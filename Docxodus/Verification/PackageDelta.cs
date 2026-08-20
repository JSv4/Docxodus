// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

namespace Docxodus.Verification;

/// <summary>Internal package-manifest change kind shared by delivery policy surfaces.</summary>
internal enum PackageDeltaChangeKind
{
    EntryAdded,
    EntryRemoved,
    EntryModified,
    RelationshipAdded,
    RelationshipRemoved,
    RelationshipModified,
}

/// <summary>
/// Policy-neutral difference between two bounded package manifests. Public delivery surfaces adapt
/// this evidence into their own versioned schema and apply their own matching/approval policy.
/// </summary>
internal sealed record PackageDeltaChange
{
    required internal PackageDeltaChangeKind Kind { get; init; }
    required internal ChangeLocation Location { get; init; }
    /// <summary>
    /// Zero-based occurrence within entries sharing a URI or relationships sharing an owner/ID.
    /// Keeping this structural key explicit lets policy adapters consume the shared delta without
    /// parsing its human-readable property path.
    /// </summary>
    required internal int Occurrence { get; init; }
    internal VerificationDigest? BeforeDigest { get; init; }
    internal VerificationDigest? AfterDigest { get; init; }
    internal string? BeforeValue { get; init; }
    internal string? AfterValue { get; init; }
}

/// <summary>A bounded package comparison. Incomplete results never expose a misleading prefix.</summary>
internal sealed record PackageDeltaResult
{
    required internal bool Complete { get; init; }
    required internal IReadOnlyList<PackageDeltaChange> Changes { get; init; }
}

/// <summary>
/// Exact entry/content-type/relationship comparison over manifests produced by the bounded package
/// inspector. It performs no delivery-policy matching and exposes no public API.
/// </summary>
internal static class PackageDelta
{
    internal static PackageDeltaResult Compare(
        PackageManifest baseline,
        PackageManifest deliverable,
        int maximumChanges)
    {
        if (maximumChanges <= 0)
            throw new ArgumentOutOfRangeException(nameof(maximumChanges));
        var changes = new List<PackageDeltaChange>();
        if (!CompareEntries(baseline, deliverable, changes, maximumChanges)
            || !CompareRelationships(baseline, deliverable, changes, maximumChanges))
            return new PackageDeltaResult
            {
                Complete = false,
                Changes = Array.Empty<PackageDeltaChange>(),
            };
        return new PackageDeltaResult
        {
            Complete = true,
            Changes = changes
                .OrderBy(change => (int)change.Kind)
                .ThenBy(change => LocationKey(change.Location), StringComparer.Ordinal)
                .ThenBy(change => change.BeforeValue, StringComparer.Ordinal)
                .ThenBy(change => change.AfterValue, StringComparer.Ordinal)
                .ToArray(),
        };
    }

    private static bool CompareEntries(
        PackageManifest baseline,
        PackageManifest deliverable,
        ICollection<PackageDeltaChange> changes,
        int maximumChanges)
    {
        var before = baseline.Entries.ToDictionary(entry => (entry.Uri, entry.Occurrence));
        var after = deliverable.Entries.ToDictionary(entry => (entry.Uri, entry.Occurrence));
        foreach (var key in before.Keys.Concat(after.Keys).Distinct()
                     .OrderBy(key => key.Uri, StringComparer.Ordinal)
                     .ThenBy(key => key.Occurrence))
        {
            before.TryGetValue(key, out var left);
            after.TryGetValue(key, out var right);
            if (left is not null && right is not null
                && string.Equals(left.ContentType, right.ContentType, StringComparison.Ordinal)
                && DigestEquals(left.RawBytesDigest, right.RawBytesDigest)
                && DigestEquals(left.NormalizedXmlDigest, right.NormalizedXmlDigest))
                continue;

            var kind = left is null
                ? PackageDeltaChangeKind.EntryAdded
                : right is null
                    ? PackageDeltaChangeKind.EntryRemoved
                    : PackageDeltaChangeKind.EntryModified;
            changes.Add(new PackageDeltaChange
            {
                Kind = kind,
                Occurrence = key.Occurrence,
                Location = new ChangeLocation
                {
                    EntryUri = key.Uri,
                    PropertyPath = $"entries[{key.Occurrence}]",
                },
                BeforeDigest = left?.RawBytesDigest,
                AfterDigest = right?.RawBytesDigest,
                BeforeValue = EntryValue(left),
                AfterValue = EntryValue(right),
            });
            if (changes.Count > maximumChanges) return false;
        }
        return true;
    }

    private static bool CompareRelationships(
        PackageManifest baseline,
        PackageManifest deliverable,
        ICollection<PackageDeltaChange> changes,
        int maximumChanges)
    {
        var before = IndexRelationships(baseline.Relationships);
        var after = IndexRelationships(deliverable.Relationships);
        foreach (var key in before.Keys.Concat(after.Keys).Distinct()
                     .OrderBy(key => key.OwnerUri, StringComparer.Ordinal)
                     .ThenBy(key => key.Id, StringComparer.Ordinal)
                     .ThenBy(key => key.Occurrence))
        {
            before.TryGetValue(key, out var left);
            after.TryGetValue(key, out var right);
            var beforeValue = RelationshipValue(left);
            var afterValue = RelationshipValue(right);
            if (string.Equals(beforeValue, afterValue, StringComparison.Ordinal))
                continue;

            var kind = left is null
                ? PackageDeltaChangeKind.RelationshipAdded
                : right is null
                    ? PackageDeltaChangeKind.RelationshipRemoved
                    : PackageDeltaChangeKind.RelationshipModified;
            changes.Add(new PackageDeltaChange
            {
                Kind = kind,
                Occurrence = key.Occurrence,
                Location = new ChangeLocation
                {
                    OwnerUri = key.OwnerUri,
                    RelationshipId = key.Id,
                    TargetUri = right?.ResolvedTargetUri ?? right?.Target
                        ?? left?.ResolvedTargetUri ?? left?.Target,
                    PropertyPath = $"relationships[{key.Occurrence}]",
                },
                BeforeValue = beforeValue,
                AfterValue = afterValue,
            });
            if (changes.Count > maximumChanges) return false;
        }
        return true;
    }

    private static Dictionary<(string OwnerUri, string Id, int Occurrence), PackageRelationship>
        IndexRelationships(IEnumerable<PackageRelationship> relationships)
    {
        var indexed = new Dictionary<(string OwnerUri, string Id, int Occurrence), PackageRelationship>();
        foreach (var group in relationships.GroupBy(
                     relationship => (relationship.OwnerUri, relationship.Id)))
        {
            int occurrence = 0;
            foreach (var relationship in group.OrderBy(RelationshipValue, StringComparer.Ordinal))
                indexed.Add((group.Key.OwnerUri, group.Key.Id, occurrence++), relationship);
        }
        return indexed;
    }

    private static string? EntryValue(PackageManifestEntry? entry) => entry is null
        ? null
        : string.Join("\u001f", new[]
        {
            entry.ContentType ?? string.Empty,
            entry.RawBytesDigest?.Value ?? string.Empty,
            entry.NormalizedXmlDigest?.Value ?? string.Empty,
        });

    private static string? RelationshipValue(PackageRelationship? relationship) =>
        relationship is null
            ? null
            : string.Join("\u001f", new[]
            {
                relationship.Type,
                relationship.Target,
                relationship.TargetMode,
                relationship.ResolvedTargetUri ?? string.Empty,
                relationship.IsTargetPresent?.ToString() ?? string.Empty,
            });

    private static string LocationKey(ChangeLocation location) => string.Join("\u001f", new[]
    {
        location.EntryUri ?? string.Empty,
        location.OwnerUri ?? string.Empty,
        location.RelationshipId ?? string.Empty,
        location.TargetUri ?? string.Empty,
        location.PropertyPath ?? string.Empty,
    });

    // Mutual absence is equality here (an entry legitimately has no XML digest);
    // comparison is the shared strict ordinal one — internally generated digests are
    // always lower-case, so the previous case-insensitivity was dead tolerance.
    private static bool DigestEquals(VerificationDigest? left, VerificationDigest? right) =>
        DeliveryReceiptValidation.OptionalDigestEquals(left, right);
}
