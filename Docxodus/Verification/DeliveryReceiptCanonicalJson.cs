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
    private static readonly DeliveryReceiptJsonContext SerializerContext =
        CreateSerializerContext();

    internal static DeliveryReceiptJsonContext JsonContext => SerializerContext;

    public static byte[] SerializeCanonical(JsonElement value) =>
        Canonicalize(Encoding.UTF8.GetBytes(value.GetRawText()));

    public static byte[] SerializeCanonicalBounded(
        JsonElement value,
        DeliveryReceiptLimits limits,
        int maximumBytes,
        string limitCode)
        => CanonicalizeBounded(
            Encoding.UTF8.GetBytes(value.GetRawText()), limits, maximumBytes, limitCode);

    internal static byte[] SerializeCanonicalBounded<T>(
        T value,
        System.Text.Json.Serialization.Metadata.JsonTypeInfo<T> typeInfo,
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
            JsonSerializer.Serialize(writer, value, typeInfo);
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
        {
            DeliveryReceiptResourceBudget.Bytes(json.Length, maximumBytes, limitCode, "JSON input");
            ScanBoundedJson(json, limits, limitCode);
        }
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

    private static void ScanBoundedJson(
        ReadOnlySpan<byte> json,
        DeliveryReceiptLimits limits,
        string limitCode)
    {
        var reader = new Utf8JsonReader(json, new JsonReaderOptions
        {
            AllowTrailingCommas = false,
            CommentHandling = JsonCommentHandling.Disallow,
            MaxDepth = limits.MaxJsonDepth,
        });
        var containers = new List<JsonScanContainer>();
        long aggregateItems = 0;

        void CountArrayValue()
        {
            if (containers.Count == 0 || containers[^1].IsObject)
                return;
            CountItem(containers[^1]);
        }

        void CountItem(JsonScanContainer container)
        {
            container.ItemCount++;
            if (container.ItemCount > limits.MaxCollectionItems)
            {
                throw new DeliveryReceiptValidationException(
                    limitCode, "JSON collection exceeds the per-collection item limit.");
            }
            aggregateItems = checked(aggregateItems + 1);
            if (aggregateItems > limits.MaxCollectionItems)
            {
                throw new DeliveryReceiptValidationException(
                    limitCode, "JSON collections exceed the aggregate item limit.");
            }
        }

        while (reader.Read())
        {
            switch (reader.TokenType)
            {
                case JsonTokenType.StartObject:
                    CountArrayValue();
                    containers.Add(new JsonScanContainer(isObject: true));
                    break;
                case JsonTokenType.StartArray:
                    CountArrayValue();
                    containers.Add(new JsonScanContainer(isObject: false));
                    break;
                case JsonTokenType.EndObject:
                case JsonTokenType.EndArray:
                    if (containers.Count == 0)
                        throw new JsonException("JSON container is unbalanced.");
                    containers.RemoveAt(containers.Count - 1);
                    break;
                case JsonTokenType.PropertyName:
                {
                    if (containers.Count == 0 || !containers[^1].IsObject)
                        throw new JsonException("JSON property is outside an object.");
                    var container = containers[^1];
                    CountItem(container);
                    var name = ReadBoundedJsonString(
                        ref reader, limits, "JSON property name", limitCode);
                    if (!container.PropertyNames!.Add(name))
                        throw new JsonException($"Duplicate JSON property '{name}' is not canonical.");
                    break;
                }
                case JsonTokenType.String:
                    CountArrayValue();
                    _ = ReadBoundedJsonString(
                        ref reader, limits, "JSON string", limitCode);
                    break;
                case JsonTokenType.Number:
                case JsonTokenType.True:
                case JsonTokenType.False:
                case JsonTokenType.Null:
                    CountArrayValue();
                    break;
                default:
                    throw new JsonException($"Unsupported JSON token {reader.TokenType}.");
            }
        }
        if (containers.Count != 0)
            throw new JsonException("JSON container is unbalanced.");
    }

    private static string ReadBoundedJsonString(
        ref Utf8JsonReader reader,
        DeliveryReceiptLimits limits,
        string name,
        string limitCode)
    {
        long maximumEncodedLength = checked((long)limits.MaxStringLength * 6);
        if (reader.ValueSpan.Length > maximumEncodedLength)
        {
            throw new DeliveryReceiptValidationException(
                limitCode, $"{name} exceeds the string-length limit.");
        }
        var value = reader.GetString()!;
        CheckJsonString(limits, value, name, limitCode);
        return value;
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

        using var document = JsonDocument.Parse(canonical, new JsonDocumentOptions
        {
            MaxDepth = DeliveryReceiptLimits.MaxAllowedJsonDepth,
        });
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
                WriteCanonicalNumber(writer, value);
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

    /// <summary>
    /// Require every known field emitted by the typed v1 model to be present with its canonical
    /// value while allowing digest-covered future optional object properties.
    /// </summary>
    public static bool ContainsCanonicalKnownProjection(
        JsonElement supplied,
        JsonElement known)
    {
        if (supplied.ValueKind != known.ValueKind)
            return false;
        switch (known.ValueKind)
        {
            case JsonValueKind.Object:
                foreach (var property in known.EnumerateObject())
                {
                    if (!supplied.TryGetProperty(property.Name, out var suppliedValue)
                        || !ContainsCanonicalKnownProjection(suppliedValue, property.Value))
                    {
                        return false;
                    }
                }
                return true;
            case JsonValueKind.Array:
            {
                var suppliedItems = supplied.EnumerateArray().ToArray();
                var knownItems = known.EnumerateArray().ToArray();
                return suppliedItems.Length == knownItems.Length
                    && suppliedItems.Zip(knownItems).All(pair =>
                        ContainsCanonicalKnownProjection(pair.First, pair.Second));
            }
            case JsonValueKind.String:
                return string.Equals(
                    supplied.GetString(), known.GetString(), StringComparison.Ordinal);
            case JsonValueKind.Number:
                return string.Equals(
                    supplied.GetRawText(), known.GetRawText(), StringComparison.Ordinal);
            case JsonValueKind.True:
            case JsonValueKind.False:
            case JsonValueKind.Null:
                return true;
            default:
                return false;
        }
    }

    public static bool HasOnlyKnownProperties(JsonElement supplied, JsonElement known)
    {
        if (supplied.ValueKind != known.ValueKind)
            return false;
        if (known.ValueKind == JsonValueKind.Object)
        {
            foreach (var property in supplied.EnumerateObject())
            {
                if (!known.TryGetProperty(property.Name, out var knownValue)
                    || !HasOnlyKnownProperties(property.Value, knownValue))
                {
                    return false;
                }
            }
            return true;
        }
        if (known.ValueKind == JsonValueKind.Array)
        {
            var suppliedItems = supplied.EnumerateArray().ToArray();
            var knownItems = known.EnumerateArray().ToArray();
            return suppliedItems.Length == knownItems.Length
                && suppliedItems.Zip(knownItems).All(pair =>
                    HasOnlyKnownProperties(pair.First, pair.Second));
        }
        return true;
    }

    private static DeliveryReceiptJsonContext CreateSerializerContext()
    {
        var options = new JsonSerializerOptions
        {
            DefaultIgnoreCondition = JsonIgnoreCondition.Never,
            Encoder = JavaScriptEncoder.Default,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            MaxDepth = 128,
            WriteIndented = false,
        };
        options.Converters.Add(new JsonStringEnumConverter<DeliveryReceiptPrivacyProfile>(
            JsonNamingPolicy.CamelCase, allowIntegerValues: false));
        options.Converters.Add(new JsonStringEnumConverter<DeliveryTransactionStatus>(
            JsonNamingPolicy.CamelCase, allowIntegerValues: false));
        options.Converters.Add(new JsonStringEnumConverter<DeliveryOperationExecutionStatus>(
            JsonNamingPolicy.CamelCase, allowIntegerValues: false));
        options.Converters.Add(new JsonStringEnumConverter<DeliveryLineageAction>(
            JsonNamingPolicy.CamelCase, allowIntegerValues: false));
        options.Converters.Add(new JsonStringEnumConverter<DeliveryChangeDisposition>(
            JsonNamingPolicy.CamelCase, allowIntegerValues: false));
        options.Converters.Add(new JsonStringEnumConverter<DeliveryPackageChangeKind>(
            JsonNamingPolicy.CamelCase, allowIntegerValues: false));
        options.Converters.Add(new JsonStringEnumConverter<DeliveryArtifactRole>(
            JsonNamingPolicy.CamelCase, allowIntegerValues: false));
        options.Converters.Add(new JsonStringEnumConverter<DeliveryArtifactAvailability>(
            JsonNamingPolicy.CamelCase, allowIntegerValues: false));
        options.Converters.Add(new JsonStringEnumConverter<DeliveryEvidenceKind>(
            JsonNamingPolicy.CamelCase, allowIntegerValues: false));
        options.Converters.Add(new JsonStringEnumConverter<DeliverySemanticComparisonScope>(
            JsonNamingPolicy.CamelCase, allowIntegerValues: false));
        options.Converters.Add(new JsonStringEnumConverter<DeliveryAuthoredEntityKind>(
            JsonNamingPolicy.CamelCase, allowIntegerValues: false));
        options.Converters.Add(new JsonStringEnumConverter<DeliveryObjectChangeKind>(
            JsonNamingPolicy.CamelCase, allowIntegerValues: false));
        options.Converters.Add(new JsonStringEnumConverter<RevisionFamily>(
            JsonNamingPolicy.CamelCase, allowIntegerValues: false));
        options.Converters.Add(new JsonStringEnumConverter<RevisionResolutionStatus>(
            JsonNamingPolicy.CamelCase, allowIntegerValues: false));
        options.Converters.Add(new JsonStringEnumConverter<MutationBatchMode>(
            JsonNamingPolicy.CamelCase, allowIntegerValues: false));
        options.Converters.Add(new JsonStringEnumConverter<PageMapStory>(
            JsonNamingPolicy.CamelCase, allowIntegerValues: false));
        return new DeliveryReceiptJsonContext(options);
    }

    private static void WriteCanonicalNumber(Utf8JsonWriter writer, JsonElement value)
    {
        // V1 deliberately preserves the exact valid JSON numeric token. This avoids lossy
        // double round-tripping for arbitrary future extension fields and is part of the
        // hash-addressed wire contract.
        writer.WriteRawValue(value.GetRawText(), skipInputValidation: false);
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
            && (text.Length > limits.MaxStringLength
                || Encoding.UTF8.GetByteCount(text) > limits.MaxStringLength))
        {
            throw new DeliveryReceiptValidationException(
                limitCode, $"{name} exceeds the string-length limit.");
        }
    }

    private sealed class JsonScanContainer
    {
        public JsonScanContainer(bool isObject)
        {
            IsObject = isObject;
            PropertyNames = isObject ? new HashSet<string>(StringComparer.Ordinal) : null;
        }

        public bool IsObject { get; }

        public HashSet<string>? PropertyNames { get; }

        public int ItemCount { get; set; }
    }
}

internal static class DeliveryReceiptValidation
{
    public const string Sha256Algorithm = "SHA-256";
    public const long MaxPortableInteger = 9_007_199_254_740_991;

    public static void ValidatePortableNonNegativeInteger(
        long value,
        string code,
        string name)
    {
        if (value < 0 || value > MaxPortableInteger)
        {
            throw new DeliveryReceiptValidationException(
                code,
                $"{name} must be between 0 and {Invariant(MaxPortableInteger)} inclusive.");
        }
    }

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

    public static string RequireOpcMainDocumentUri(string? value, string name)
    {
        if (string.IsNullOrWhiteSpace(value)
            || value.Length > 4096
            || value[0] != '/'
            || value.Contains('\\'))
        {
            throw new DeliveryReceiptValidationException(
                "not_wordprocessing_package",
                $"{name} must be an absolute OPC part URI.");
        }
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
        var limits = new DeliveryReceiptLimits().ValidateAndClone();
        DeliveryReceiptResourceValidator.ValidatePayload(payload, limits);
        return SerializePayload(payload, limits);
    }

    internal static byte[] SerializePayload(
        DeliveryChangeReceiptPayload payload,
        DeliveryReceiptLimits limits)
    {
        ArgumentNullException.ThrowIfNull(payload);
        return DeliveryReceiptCanonicalJson.SerializeCanonicalBounded(
            payload,
            DeliveryReceiptCanonicalJson.JsonContext.DeliveryChangeReceiptPayload,
            limits,
            limits.MaxReceiptJsonBytes,
            "receipt_resource_limit");
    }

    public static byte[] Serialize(DeliveryChangeReceipt receipt, bool indented = false)
        => Serialize(receipt, new DeliveryReceiptLimits(), indented);

    public static byte[] Serialize(
        DeliveryChangeReceipt receipt,
        DeliveryReceiptLimits limits,
        bool indented = false)
    {
        ArgumentNullException.ThrowIfNull(receipt);
        ArgumentNullException.ThrowIfNull(limits);
        var validatedLimits = limits.ValidateAndClone();
        DeliveryReceiptResourceValidator.ValidatePayload(receipt.Payload, validatedLimits);
        var payload = SerializePayload(receipt.Payload, validatedLimits);
        DeliveryReceiptValidation.ValidateDigest(receipt.ReceiptDigest, "receipt digest");

        using var payloadDocument = JsonDocument.Parse(payload, new JsonDocumentOptions
        {
            MaxDepth = DeliveryReceiptLimits.MaxAllowedJsonDepth,
        });
        using var stream = new DeliveryReceiptBoundedMemoryStream(
            validatedLimits.MaxReceiptJsonBytes,
            "receipt_resource_limit",
            "Receipt JSON");
        using (var writer = new Utf8JsonWriter(stream, new JsonWriterOptions
        {
            Encoder = JavaScriptEncoder.Default,
            Indented = false,
            MaxDepth = 128,
        }))
        {
            writer.WriteStartObject();
            writer.WritePropertyName("payload");
            payloadDocument.RootElement.WriteTo(writer);
            writer.WriteStartObject("receiptDigest");
            writer.WriteString("algorithm", receipt.ReceiptDigest.Algorithm);
            writer.WriteString("value", receipt.ReceiptDigest.Value);
            writer.WriteEndObject();
            writer.WriteEndObject();
        }
        var canonical = stream.ToArray();
        if (!indented)
            return canonical;

        using var canonicalDocument = JsonDocument.Parse(canonical, new JsonDocumentOptions
        {
            MaxDepth = DeliveryReceiptLimits.MaxAllowedJsonDepth,
        });
        using var indentedStream = new DeliveryReceiptBoundedMemoryStream(
            validatedLimits.MaxReceiptJsonBytes,
            "receipt_resource_limit",
            "Indented receipt JSON");
        using (var writer = new Utf8JsonWriter(indentedStream, new JsonWriterOptions
        {
            Encoder = JavaScriptEncoder.Default,
            Indented = true,
            MaxDepth = 128,
        }))
        {
            canonicalDocument.RootElement.WriteTo(writer);
        }
        return indentedStream.ToArray();
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
                DeliveryReceiptCanonicalJson.JsonContext.DeliveryChangeReceipt,
                limits,
                limits.MaxReceiptJsonBytes,
                "receipt_resource_limit").LongLength,
            limits.MaxReceiptJsonBytes,
            "receipt_resource_limit",
            "Receipt JSON");
        return receipt;
    }
}
