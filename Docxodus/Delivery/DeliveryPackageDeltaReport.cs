// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using Docxodus.Verification;

namespace Docxodus.Delivery;

/// <summary>Stable package change kinds in the delivery-bundle delta artifact.</summary>
public enum DeliveryPackageDeltaChangeKind
{
    EntryAdded,
    EntryRemoved,
    EntryModified,
    RelationshipAdded,
    RelationshipRemoved,
    RelationshipModified,
}

/// <summary>One complete-package identity recorded by a package-delta artifact.</summary>
public sealed record DeliveryPackageDeltaDocumentIdentity
{
    required public VerificationDigest RawPackageBytesDigest { get; init; }
    public VerificationDigest? OrderedOpcContentDigest { get; init; }
    public VerificationDigest? NormalizedSemanticDigest { get; init; }
}

/// <summary>One deterministic package-entry or relationship change.</summary>
public sealed record DeliveryPackageDeltaChange
{
    required public string ChangeId { get; init; }
    required public DeliveryPackageDeltaChangeKind Kind { get; init; }
    required public ChangeLocation Location { get; init; }
    required public int Occurrence { get; init; }
    public VerificationDigest? BeforeDigest { get; init; }
    public VerificationDigest? AfterDigest { get; init; }
    public string? BeforeValue { get; init; }
    public string? AfterValue { get; init; }
}

/// <summary>
/// Versioned, policy-neutral package delta used by a delivery bundle. The comparison delegates to
/// the same bounded manifest delta owner used by deliverable verification, receipts, and redline
/// proofs; this type only supplies the public artifact schema.
/// </summary>
public sealed record DeliveryPackageDeltaReport
{
    public const int DefaultMaximumChanges = 25_000;

    public const string SchemaId =
        "https://docxodus.dev/schemas/delivery/package-delta/v1";

    public string Schema { get; init; } = SchemaId;
    public int SchemaVersion { get; init; } = 1;
    required public DeliveryPackageDeltaDocumentIdentity BaselineDocument { get; init; }
    required public DeliveryPackageDeltaDocumentIdentity FinalDocument { get; init; }
    required public int ChangeCount { get; init; }
    public IReadOnlyList<DeliveryPackageDeltaChange> Changes { get; init; } =
        Array.Empty<DeliveryPackageDeltaChange>();

    public static DeliveryPackageDeltaReport Create(
        PackageManifest baseline,
        PackageManifest final,
        int maximumChanges = DefaultMaximumChanges)
    {
        ArgumentNullException.ThrowIfNull(baseline);
        ArgumentNullException.ThrowIfNull(final);
        if (maximumChanges <= 0)
            throw new ArgumentOutOfRangeException(nameof(maximumChanges));
        if (!baseline.IsValid || !final.IsValid)
        {
            throw new ArgumentException(
                "Package delta inputs must be valid bounded package manifests.");
        }

        var delta = PackageDelta.Compare(baseline, final, maximumChanges);
        if (!delta.Complete)
        {
            throw new InvalidDataException(
                "Package delta exceeds the configured complete-change limit.");
        }
        var changes = delta.Changes.Select(Project)
            .ToArray();
        return new DeliveryPackageDeltaReport
        {
            BaselineDocument = Identity(baseline),
            FinalDocument = Identity(final),
            ChangeCount = changes.Length,
            Changes = changes,
        };
    }

    public byte[] ToCanonicalUtf8Bytes() =>
        JsonSerializer.SerializeToUtf8Bytes(this, JsonOptions.Canonical);

    public string ToCanonicalJson() => Encoding.UTF8.GetString(ToCanonicalUtf8Bytes());

    public string ToJson(bool indented = true) => JsonSerializer.Serialize(
        this, indented ? JsonOptions.Indented : JsonOptions.Canonical);

    private static DeliveryPackageDeltaDocumentIdentity Identity(PackageManifest manifest) => new()
    {
        RawPackageBytesDigest = manifest.RawPackageBytesDigest,
        OrderedOpcContentDigest = manifest.OrderedOpcContentDigest,
        NormalizedSemanticDigest = manifest.NormalizedSemanticDigest,
    };

    private static DeliveryPackageDeltaChange Project(PackageDeltaChange change)
    {
        var kind = change.Kind switch
        {
            PackageDeltaChangeKind.EntryAdded => DeliveryPackageDeltaChangeKind.EntryAdded,
            PackageDeltaChangeKind.EntryRemoved => DeliveryPackageDeltaChangeKind.EntryRemoved,
            PackageDeltaChangeKind.EntryModified => DeliveryPackageDeltaChangeKind.EntryModified,
            PackageDeltaChangeKind.RelationshipAdded =>
                DeliveryPackageDeltaChangeKind.RelationshipAdded,
            PackageDeltaChangeKind.RelationshipRemoved =>
                DeliveryPackageDeltaChangeKind.RelationshipRemoved,
            PackageDeltaChangeKind.RelationshipModified =>
                DeliveryPackageDeltaChangeKind.RelationshipModified,
            _ => throw new ArgumentOutOfRangeException(nameof(change), change.Kind, null),
        };
        var identity = string.Join("\u001f", new[]
        {
            kind.ToString(),
            change.Location.EntryUri ?? string.Empty,
            change.Location.OwnerUri ?? string.Empty,
            change.Location.RelationshipId ?? string.Empty,
            change.Location.TargetUri ?? string.Empty,
            change.Location.PropertyPath ?? string.Empty,
            change.Occurrence.ToString(System.Globalization.CultureInfo.InvariantCulture),
            change.BeforeDigest?.Algorithm ?? string.Empty,
            change.BeforeDigest?.Value ?? string.Empty,
            change.AfterDigest?.Algorithm ?? string.Empty,
            change.AfterDigest?.Value ?? string.Empty,
            change.BeforeValue ?? string.Empty,
            change.AfterValue ?? string.Empty,
        });
        return new DeliveryPackageDeltaChange
        {
            ChangeId = "pkg-" + Convert.ToHexString(
                    SHA256.HashData(Encoding.UTF8.GetBytes(identity)))
                .ToLowerInvariant()[..20],
            Kind = kind,
            Location = change.Location,
            Occurrence = change.Occurrence,
            BeforeDigest = change.BeforeDigest,
            AfterDigest = change.AfterDigest,
            BeforeValue = change.BeforeValue,
            AfterValue = change.AfterValue,
        };
    }

    private static class JsonOptions
    {
        internal static readonly JsonSerializerOptions Canonical = Create(indented: false);
        internal static readonly JsonSerializerOptions Indented = Create(indented: true);

        private static JsonSerializerOptions Create(bool indented)
        {
            var options = new JsonSerializerOptions
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                WriteIndented = indented,
            };
            options.Converters.Add(new JsonStringEnumConverter(JsonNamingPolicy.CamelCase));
            return options;
        }
    }
}
