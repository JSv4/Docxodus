// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace Docxodus.Verification;

internal static class DeliveryReceiptCanonicalJson
{
    private static readonly JsonSerializerOptions SerializerOptions = new()
    {
        DefaultIgnoreCondition = JsonIgnoreCondition.Never,
        Encoder = JavaScriptEncoder.Default,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        MaxDepth = 128,
        WriteIndented = false,
        Converters =
        {
            new JsonStringEnumConverter(JsonNamingPolicy.CamelCase, allowIntegerValues: false),
        },
    };

    public static JsonSerializerOptions JsonOptions => SerializerOptions;

    public static byte[] SerializeCanonical<T>(T value)
    {
        var json = JsonSerializer.SerializeToUtf8Bytes(value, SerializerOptions);
        return Canonicalize(json);
    }

    public static byte[] SerializeCanonicalBounded<T>(
        T value,
        DeliveryReceiptLimits limits,
        int maximumBytes,
        string limitCode)
    {
        ArgumentNullException.ThrowIfNull(limits);
        using var stream = new DeliveryReceiptBoundedMemoryStream(
            maximumBytes, limitCode, "Serialized JSON");
        using (var writer = new Utf8JsonWriter(stream, new JsonWriterOptions
        {
            Encoder = JavaScriptEncoder.Default,
            Indented = false,
            MaxDepth = limits.MaxJsonDepth,
        }))
        {
            var options = new JsonSerializerOptions(SerializerOptions)
            {
                MaxDepth = limits.MaxJsonDepth,
            };
            JsonSerializer.Serialize(writer, value, options);
        }
        var json = stream.ToArray();
        return CanonicalizeBounded(json, limits, maximumBytes, limitCode);
    }

    public static byte[] Canonicalize(ReadOnlySpan<byte> json)
        => CanonicalizeCore(json, null, int.MaxValue, "receipt_resource_limit");

    public static byte[] CanonicalizeBounded(
        ReadOnlySpan<byte> json,
        DeliveryReceiptLimits limits,
        int maximumBytes,
        string limitCode)
    {
        ArgumentNullException.ThrowIfNull(limits);
        DeliveryReceiptResourceBudget.Bytes(json.Length, maximumBytes, limitCode, "JSON input");
        return CanonicalizeCore(json, limits, maximumBytes, limitCode);
    }

    private static byte[] CanonicalizeCore(
        ReadOnlySpan<byte> json,
        DeliveryReceiptLimits? limits,
        int maximumBytes,
        string limitCode)
    {
        if (limits is not null)
            DeliveryReceiptResourceBudget.Bytes(json.Length, maximumBytes, limitCode, "JSON input");
        using var document = JsonDocument.Parse(json.ToArray(), new JsonDocumentOptions
        {
            AllowTrailingCommas = false,
            CommentHandling = JsonCommentHandling.Disallow,
            MaxDepth = limits?.MaxJsonDepth ?? 128,
        });
        using MemoryStream stream = limits is null
            ? new MemoryStream()
            : new DeliveryReceiptBoundedMemoryStream(
                maximumBytes, limitCode, "Canonical JSON");
        long collectionItems = 0;
        using (var writer = new Utf8JsonWriter(stream, new JsonWriterOptions
        {
            Encoder = JavaScriptEncoder.Default,
            Indented = false,
            MaxDepth = limits?.MaxJsonDepth ?? 128,
        }))
        {
            WriteCanonical(
                writer, document.RootElement, limits, ref collectionItems, limitCode);
        }
        var canonical = stream.ToArray();
        if (limits is not null)
        {
            DeliveryReceiptResourceBudget.Bytes(
                canonical.LongLength, maximumBytes, limitCode, "Canonical JSON");
        }
        return canonical;
    }

    public static JsonElement ParseCanonicalObject(string json, string parameterName)
    {
        if (json is null)
            throw new ArgumentNullException(parameterName);
        byte[] canonical;
        try
        {
            canonical = Canonicalize(Encoding.UTF8.GetBytes(json));
        }
        catch (JsonException ex)
        {
            throw new DeliveryReceiptValidationException(
                "invalid_operation_arguments", $"{parameterName} is not valid JSON: {ex.Message}");
        }

        using var document = JsonDocument.Parse(canonical);
        if (document.RootElement.ValueKind != JsonValueKind.Object)
        {
            throw new DeliveryReceiptValidationException(
                "invalid_operation_arguments", $"{parameterName} must be a JSON object.");
        }
        return document.RootElement.Clone();
    }

    public static VerificationDigest Digest(ReadOnlySpan<byte> bytes) => new()
    {
        Algorithm = DeliveryReceiptValidation.Sha256Algorithm,
        Value = Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant(),
    };

    public static VerificationDigest DigestText(string value) =>
        Digest(Encoding.UTF8.GetBytes(value));

    public static bool FixedTimeEquals(VerificationDigest expected, ReadOnlySpan<byte> bytes)
    {
        DeliveryReceiptValidation.ValidateDigest(expected, "expected digest");
        byte[] expectedBytes;
        try { expectedBytes = Convert.FromHexString(expected.Value); }
        catch (FormatException) { return false; }
        var actual = SHA256.HashData(bytes);
        return CryptographicOperations.FixedTimeEquals(expectedBytes, actual);
    }

    public static string DigestToken(ReadOnlySpan<byte> bytes) =>
        "sha256:" + Digest(bytes).Value;

    private static void WriteCanonical(
        Utf8JsonWriter writer,
        JsonElement value,
        DeliveryReceiptLimits? limits,
        ref long collectionItems,
        string limitCode)
    {
        switch (value.ValueKind)
        {
            case JsonValueKind.Object:
            {
                writer.WriteStartObject();
                var properties = new List<JsonProperty>();
                foreach (var property in value.EnumerateObject())
                {
                    AddJsonItems(limits, 1, ref collectionItems, limitCode);
                    CheckJsonString(
                        limits, property.Name, "JSON property name", limitCode);
                    properties.Add(property);
                }
                properties.Sort(static (left, right) =>
                    string.CompareOrdinal(left.Name, right.Name));
                for (int i = 1; i < properties.Count; i++)
                {
                    if (string.Equals(properties[i - 1].Name, properties[i].Name,
                            StringComparison.Ordinal))
                    {
                        throw new JsonException(
                            $"Duplicate JSON property '{properties[i].Name}' is not canonical.");
                    }
                }
                foreach (var property in properties)
                {
                    writer.WritePropertyName(property.Name);
                    WriteCanonical(
                        writer, property.Value, limits, ref collectionItems, limitCode);
                }
                writer.WriteEndObject();
                break;
            }
            case JsonValueKind.Array:
                writer.WriteStartArray();
                foreach (var item in value.EnumerateArray())
                {
                    AddJsonItems(limits, 1, ref collectionItems, limitCode);
                    WriteCanonical(writer, item, limits, ref collectionItems, limitCode);
                }
                writer.WriteEndArray();
                break;
            case JsonValueKind.String:
                var stringValue = value.GetString();
                CheckJsonString(limits, stringValue, "JSON string", limitCode);
                writer.WriteStringValue(stringValue);
                break;
            case JsonValueKind.Number:
                // Numeric spelling is part of v1 for arbitrary operation/evidence extensions.
                // Core receipt numbers are emitted by System.Text.Json using invariant formatting.
                writer.WriteRawValue(value.GetRawText(), skipInputValidation: false);
                break;
            case JsonValueKind.True:
                writer.WriteBooleanValue(true);
                break;
            case JsonValueKind.False:
                writer.WriteBooleanValue(false);
                break;
            case JsonValueKind.Null:
                writer.WriteNullValue();
                break;
            default:
                throw new JsonException($"Unsupported JSON token {value.ValueKind}.");
        }

    }

    private static void AddJsonItems(
        DeliveryReceiptLimits? limits,
        int count,
        ref long collectionItems,
        string limitCode)
    {
        if (limits is null)
            return;
        if (count > limits.MaxCollectionItems)
            throw new DeliveryReceiptValidationException(
                limitCode, "JSON collection exceeds the per-collection item limit.");
        collectionItems = checked(collectionItems + count);
        if (collectionItems > limits.MaxCollectionItems)
            throw new DeliveryReceiptValidationException(
                limitCode, "JSON collections exceed the aggregate item limit.");
    }

    private static void CheckJsonString(
        DeliveryReceiptLimits? limits,
        string? text,
        string name,
        string limitCode)
    {
        if (limits is not null && text is not null
            && text.Length > limits.MaxStringLength)
        {
            throw new DeliveryReceiptValidationException(
                limitCode, $"{name} exceeds the string-length limit.");
        }
    }
}

internal static class DeliveryReceiptValidation
{
    public const string Sha256Algorithm = "SHA-256";

    public static void ValidateDigest(VerificationDigest? digest, string name)
    {
        if (digest is null)
            throw new DeliveryReceiptValidationException("missing_digest", $"{name} is required.");
        if (!string.Equals(digest.Algorithm, Sha256Algorithm, StringComparison.Ordinal))
        {
            throw new DeliveryReceiptValidationException(
                "unsupported_digest_algorithm", $"{name} must use {Sha256Algorithm}.");
        }
        if (digest.Value is null || digest.Value.Length != 64
            || digest.Value.Any(c => !((c >= '0' && c <= '9') || (c >= 'a' && c <= 'f'))))
        {
            throw new DeliveryReceiptValidationException(
                "invalid_digest", $"{name} must be 64 lower-case hexadecimal characters.");
        }
    }

    public static void ValidateOptionalDigest(VerificationDigest? digest, string name)
    {
        if (digest is not null)
            ValidateDigest(digest, name);
    }

    public static VerificationDigest CloneDigest(VerificationDigest digest)
    {
        ValidateDigest(digest, "digest");
        return new VerificationDigest { Algorithm = digest.Algorithm, Value = digest.Value };
    }

    public static VerificationDigest? CloneOptionalDigest(VerificationDigest? digest) =>
        digest is null ? null : CloneDigest(digest);

    public static string RequireNonBlank(string? value, string name, int maxLength = 2048)
    {
        if (string.IsNullOrWhiteSpace(value))
            throw new DeliveryReceiptValidationException("missing_value", $"{name} is required.");
        if (value.Length > maxLength)
            throw new DeliveryReceiptValidationException("value_too_long", $"{name} is too long.");
        return value;
    }

    public static string? NormalizeRelativePath(string? path)
    {
        if (path is null)
            return null;
        RequireNonBlank(path, "artifact relative path", 4096);
        if (path[0] is '/' or '\\'
            || (path.Length >= 2
                && ((path[0] >= 'A' && path[0] <= 'Z')
                    || (path[0] >= 'a' && path[0] <= 'z'))
                && path[1] == ':'))
        {
            throw new DeliveryReceiptValidationException(
                "unsafe_artifact_path", "Artifact display paths must be portable relative paths.");
        }
        var normalized = path.Replace('\\', '/');
        var segments = normalized.Split('/');
        if (segments.Any(segment => segment is "" or "." or ".."))
        {
            throw new DeliveryReceiptValidationException(
                "unsafe_artifact_path", "Artifact display paths cannot contain empty, dot, or parent segments.");
        }
        return string.Join('/', segments);
    }

    public static bool DigestEquals(VerificationDigest? left, VerificationDigest? right) =>
        left is not null && right is not null
        && string.Equals(left.Algorithm, right.Algorithm, StringComparison.Ordinal)
        && string.Equals(left.Value, right.Value, StringComparison.Ordinal);

    public static string Invariant(long value) => value.ToString(CultureInfo.InvariantCulture);
}

/// <summary>Canonical serializer for the delivery-receipt v1 envelope.</summary>
public static class DeliveryChangeReceiptSerializer
{
    public static byte[] SerializePayload(DeliveryChangeReceiptPayload payload)
    {
        ArgumentNullException.ThrowIfNull(payload);
        return DeliveryReceiptCanonicalJson.SerializeCanonical(payload);
    }

    internal static byte[] SerializePayload(
        DeliveryChangeReceiptPayload payload,
        DeliveryReceiptLimits limits)
    {
        ArgumentNullException.ThrowIfNull(payload);
        return DeliveryReceiptCanonicalJson.SerializeCanonicalBounded(
            payload,
            limits,
            limits.MaxReceiptJsonBytes,
            "receipt_resource_limit");
    }

    public static byte[] Serialize(DeliveryChangeReceipt receipt, bool indented = false)
    {
        ArgumentNullException.ThrowIfNull(receipt);
        var payload = SerializePayload(receipt.Payload);
        DeliveryReceiptValidation.ValidateDigest(receipt.ReceiptDigest, "receipt digest");

        using var payloadDocument = JsonDocument.Parse(payload);
        var envelope = new Dictionary<string, object?>
        {
            ["payload"] = payloadDocument.RootElement.Clone(),
            ["receiptDigest"] = receipt.ReceiptDigest,
        };
        var canonical = DeliveryReceiptCanonicalJson.SerializeCanonical(envelope);
        if (!indented)
            return canonical;

        using var canonicalDocument = JsonDocument.Parse(canonical);
        return JsonSerializer.SerializeToUtf8Bytes(canonicalDocument.RootElement,
            new JsonSerializerOptions(DeliveryReceiptCanonicalJson.JsonOptions)
            {
                WriteIndented = true,
            });
    }

    internal static DeliveryChangeReceipt Create(
        DeliveryChangeReceiptPayload payload,
        DeliveryReceiptLimits limits)
    {
        var canonicalPayload = SerializePayload(payload, limits);
        var receipt = new DeliveryChangeReceipt
        {
            Payload = payload,
            ReceiptDigest = DeliveryReceiptCanonicalJson.Digest(canonicalPayload),
        };
        DeliveryReceiptResourceBudget.Bytes(
            DeliveryReceiptCanonicalJson.SerializeCanonicalBounded(
                receipt,
                limits,
                limits.MaxReceiptJsonBytes,
                "receipt_resource_limit").LongLength,
            limits.MaxReceiptJsonBytes,
            "receipt_resource_limit",
            "Receipt JSON");
        return receipt;
    }
}
