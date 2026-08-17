// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

namespace Docxodus.Verification;

/// <summary>
/// The single #456 integration point for delivery receipts. Package parsing remains owned by
/// PackageManifestGenerator; this adapter only projects the public manifest into immutable receipt
/// identities and adapts #463's policy-neutral package delta into receipt evidence. Corrected #493
/// availability semantics for optional content/semantic digests remain isolated here rather than
/// being reinterpreted by receipts.
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
        DeliveryReceiptValidation.ValidatePortableNonNegativeInteger(
            documentVersion, "invalid_document_version", "Document version");
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
        if (!string.Equals(packageKind, "opc", StringComparison.Ordinal))
        {
            throw new DeliveryReceiptValidationException(
                "not_wordprocessing_package",
                "Delivery document identities require an OPC Word main-document part.");
        }
        var mainDocumentUri = DeliveryReceiptValidation.RequireOpcMainDocumentUri(
            manifest.Facts.MainDocumentUri, "main document URI");
        return new DeliveryDocumentIdentity
        {
            DocumentVersion = documentVersion,
            PackageKind = packageKind,
            PackageManifestSchema = manifest.Schema,
            MainDocumentUri = mainDocumentUri,
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
        PackageManifest after,
        int maximumChanges)
    {
        ArgumentNullException.ThrowIfNull(before);
        ArgumentNullException.ThrowIfNull(after);
        var delta = PackageDelta.Compare(before, after, maximumChanges);
        if (!delta.Complete)
        {
            throw new DeliveryReceiptValidationException(
                "receipt_resource_limit",
                "Package comparison exceeds the receipt change-record limit.");
        }
        return delta.Changes.Select(ProjectChange).ToArray();
    }

    private static DeliveryPackageChangeObservation ProjectChange(PackageDeltaChange change) =>
        new(
            change.Kind switch
            {
                PackageDeltaChangeKind.EntryAdded => DeliveryPackageChangeKind.PartAdded,
                PackageDeltaChangeKind.EntryRemoved => DeliveryPackageChangeKind.PartRemoved,
                PackageDeltaChangeKind.EntryModified => DeliveryPackageChangeKind.PartModified,
                PackageDeltaChangeKind.RelationshipAdded =>
                    DeliveryPackageChangeKind.RelationshipAdded,
                PackageDeltaChangeKind.RelationshipRemoved =>
                    DeliveryPackageChangeKind.RelationshipRemoved,
                PackageDeltaChangeKind.RelationshipModified =>
                    DeliveryPackageChangeKind.RelationshipModified,
                _ => throw new ArgumentOutOfRangeException(nameof(change), change.Kind, null),
            },
            change.Location,
            change.BeforeValue,
            change.AfterValue);
}

internal sealed record DeliveryPackageChangeObservation(
    DeliveryPackageChangeKind Kind,
    ChangeLocation Location,
    string? Before,
    string? After);
