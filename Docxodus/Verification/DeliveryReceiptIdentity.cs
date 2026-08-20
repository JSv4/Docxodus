// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Text.Encodings.Web;
using System.Text.Json;

namespace Docxodus.Verification;

/// <summary>
/// Trim/AOT-safe writers for the small canonical projections used as receipt identities.
/// Keeping construction and portable verification on these shared writers prevents drift.
/// </summary>
internal static class DeliveryReceiptIdentity
{
    public static string PackageChangeId(
        DeliveryPackageChangeKind kind,
        ChangeLocation location,
        VerificationDigest? before,
        VerificationDigest? after)
    {
        ArgumentNullException.ThrowIfNull(location);
        return DeliveryReceiptCanonicalJson.DigestToken(Write(writer =>
        {
            writer.WriteStartObject();
            writer.WritePropertyName("after");
            WriteDigest(writer, after);
            writer.WritePropertyName("before");
            WriteDigest(writer, before);
            writer.WriteString("kind", kind.ToString());
            writer.WriteStartObject("location");
            WriteNullableString(writer, "entryUri", location.EntryUri);
            WriteNullableString(writer, "ownerUri", location.OwnerUri);
            WriteNullableString(writer, "propertyPath", location.PropertyPath);
            WriteNullableString(writer, "relationshipId", location.RelationshipId);
            WriteNullableString(writer, "targetUri", location.TargetUri);
            writer.WriteEndObject();
            writer.WriteEndObject();
        }));
    }

    public static string TransactionEntryId(
        string requestFingerprint,
        DeliveryDocumentIdentity before,
        DeliveryDocumentIdentity after,
        long baseVersion,
        long resultVersion,
        string? transactionId,
        long sequence)
    {
        ArgumentNullException.ThrowIfNull(before);
        ArgumentNullException.ThrowIfNull(after);
        return DeliveryReceiptCanonicalJson.DigestToken(Write(writer =>
        {
            writer.WriteStartObject();
            writer.WritePropertyName("afterPackage");
            WriteDigest(writer, after.RawPackageBytesDigest);
            writer.WriteNumber("baseVersion", baseVersion);
            writer.WritePropertyName("beforePackage");
            WriteDigest(writer, before.RawPackageBytesDigest);
            writer.WriteString("requestFingerprint", requestFingerprint);
            writer.WriteNumber("resultVersion", resultVersion);
            WriteNullableString(writer, "transactionId", transactionId);
            if (transactionId is null)
                writer.WriteNumber("transactionSequence", sequence);
            else
                writer.WriteNull("transactionSequence");
            writer.WriteEndObject();
        }));
    }

    public static string RequestFingerprint(
        MutationBatchMode mode,
        IReadOnlyList<DeliveryNormalizedOperation> operations,
        DeliveryReceiptLimits limits)
    {
        ArgumentNullException.ThrowIfNull(operations);
        ArgumentNullException.ThrowIfNull(limits);
        var bytes = WriteBounded(writer =>
        {
            writer.WriteStartObject();
            writer.WriteString(
                "mode", mode == MutationBatchMode.Atomic ? "atomic" : "bestEffort");
            writer.WriteStartArray("operations");
            foreach (var operation in operations)
            {
                writer.WriteStartObject();
                writer.WriteString("action", operation.Action);
                writer.WritePropertyName("arguments");
                operation.Arguments.WriteTo(writer);
                writer.WriteString("tool", operation.Tool);
                writer.WriteEndObject();
            }
            writer.WriteEndArray();
            writer.WriteEndObject();
        }, limits);
        return DeliveryReceiptCanonicalJson.DigestToken(bytes);
    }

    public static byte[] TransactionEvidence(
        DeliveryTransactionEntry entry,
        DeliveryReceiptLimits limits) =>
        DeliveryReceiptCanonicalJson.SerializeCanonicalBounded(
            entry with { Sequence = 0 },
            DeliveryReceiptCanonicalJson.JsonContext.DeliveryTransactionEntry,
            limits,
            limits.MaxReceiptJsonBytes,
            "receipt_resource_limit");

    private static byte[] Write(Action<Utf8JsonWriter> write)
    {
        using var stream = new MemoryStream();
        using (var writer = NewWriter(stream, 128))
            write(writer);
        return DeliveryReceiptCanonicalJson.Canonicalize(stream.ToArray());
    }

    private static byte[] WriteBounded(
        Action<Utf8JsonWriter> write,
        DeliveryReceiptLimits limits)
    {
        using var stream = new DeliveryReceiptBoundedMemoryStream(
            limits.MaxReceiptJsonBytes,
            "receipt_resource_limit",
            "Receipt identity JSON");
        using (var writer = NewWriter(stream, limits.MaxJsonDepth))
            write(writer);
        return DeliveryReceiptCanonicalJson.CanonicalizeBounded(
            stream.ToArray(),
            limits,
            limits.MaxReceiptJsonBytes,
            "receipt_resource_limit");
    }

    private static Utf8JsonWriter NewWriter(Stream stream, int maxDepth) => new(
        stream,
        new JsonWriterOptions
        {
            Encoder = JavaScriptEncoder.Default,
            Indented = false,
            MaxDepth = maxDepth,
        });

    private static void WriteDigest(Utf8JsonWriter writer, VerificationDigest? digest)
    {
        if (digest is null)
        {
            writer.WriteNullValue();
            return;
        }
        writer.WriteStartObject();
        writer.WriteString("algorithm", digest.Algorithm);
        writer.WriteString("value", digest.Value);
        writer.WriteEndObject();
    }

    private static void WriteNullableString(
        Utf8JsonWriter writer,
        string propertyName,
        string? value)
    {
        if (value is null)
            writer.WriteNull(propertyName);
        else
            writer.WriteString(propertyName, value);
    }
}
