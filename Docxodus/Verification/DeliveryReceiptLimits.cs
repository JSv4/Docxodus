// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;

namespace Docxodus.Verification;

/// <summary>Bounded resource policy shared by receipt construction and verification.</summary>
public sealed record DeliveryReceiptLimits
{
    public int MaxReceiptJsonBytes { get; init; } = 16 * 1024 * 1024;
    public int MaxSemanticEvidenceBytes { get; init; } = 64 * 1024 * 1024;
    public int MaxPageMapBytes { get; init; } = 64 * 1024 * 1024;
    public int MaxArtifactBytes { get; init; } = 256 * 1024 * 1024;
    public long MaxTotalArtifactBytes { get; init; } = 512L * 1024 * 1024;
    /// <summary>Hard ceiling for <see cref="MaxJsonDepth"/>; every secondary parse of
    /// already-bounded canonical JSON must allow at least this depth.</summary>
    internal const int MaxAllowedJsonDepth = 128;

    public int MaxJsonDepth { get; init; } = MaxAllowedJsonDepth;
    public int MaxCollectionItems { get; init; } = 100_000;
    public int MaxTransactions { get; init; } = 10_000;
    public int MaxOperationsPerTransaction { get; init; } = 10_000;
    public int MaxArtifacts { get; init; } = 1_024;
    public int MaxStringLength { get; init; } = 1024 * 1024;

    /// <summary>
    /// Corrected #493 inspection policy used when clean-DOCX bytes are independently reparsed.
    /// Package expansion, entry, XML, ratio, and URI limits remain owned by that contract.
    /// </summary>
    public PackageManifestOptions CleanDocxManifestOptions { get; init; } = new();

    internal DeliveryReceiptLimits ValidateAndClone()
    {
        Positive(MaxReceiptJsonBytes, nameof(MaxReceiptJsonBytes));
        Positive(MaxSemanticEvidenceBytes, nameof(MaxSemanticEvidenceBytes));
        Positive(MaxPageMapBytes, nameof(MaxPageMapBytes));
        Positive(MaxArtifactBytes, nameof(MaxArtifactBytes));
        if (MaxTotalArtifactBytes <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxTotalArtifactBytes));
        if (MaxJsonDepth is < 1 or > MaxAllowedJsonDepth)
            throw new ArgumentOutOfRangeException(nameof(MaxJsonDepth));
        Positive(MaxCollectionItems, nameof(MaxCollectionItems));
        Positive(MaxTransactions, nameof(MaxTransactions));
        Positive(MaxOperationsPerTransaction, nameof(MaxOperationsPerTransaction));
        Positive(MaxArtifacts, nameof(MaxArtifacts));
        Positive(MaxStringLength, nameof(MaxStringLength));
        ArgumentNullException.ThrowIfNull(CleanDocxManifestOptions);
        CleanDocxManifestOptions.Validate();
        return this with
        {
            CleanDocxManifestOptions = CloneManifestOptions(CleanDocxManifestOptions),
        };
    }

    private static void Positive(int value, string name)
    {
        if (value <= 0)
            throw new ArgumentOutOfRangeException(name);
    }

    private static PackageManifestOptions CloneManifestOptions(PackageManifestOptions value) => new()
    {
        MaxEntryCount = value.MaxEntryCount,
        MaxEntryUncompressedBytes = value.MaxEntryUncompressedBytes,
        MaxTotalUncompressedBytes = value.MaxTotalUncompressedBytes,
        MaxXmlPartBytes = value.MaxXmlPartBytes,
        MaxCompressionRatio = value.MaxCompressionRatio,
        MaxUriLength = value.MaxUriLength,
    };
}

/// <summary>Options for portable receipt verification.</summary>
public sealed record DeliveryReceiptVerificationOptions
{
    public DeliveryReceiptLimits Limits { get; init; } = new();

    internal DeliveryReceiptVerificationOptions ValidateAndClone()
    {
        ArgumentNullException.ThrowIfNull(Limits);
        return this with { Limits = Limits.ValidateAndClone() };
    }
}

internal sealed class DeliveryReceiptResourceBudget
{
    private readonly DeliveryReceiptLimits _limits;
    private readonly long _maximumSerializedBytes;
    private readonly string _limitCode;
    private long _collectionItems;
    private long _serializedBytes;

    public DeliveryReceiptResourceBudget(
        DeliveryReceiptLimits limits,
        long? maximumSerializedBytes = null,
        string limitCode = "receipt_resource_limit")
    {
        _limits = limits;
        _maximumSerializedBytes = maximumSerializedBytes ?? limits.MaxReceiptJsonBytes;
        _limitCode = limitCode;
    }

    public void AddItems(int count, string name)
    {
        if (count < 0 || count > _limits.MaxCollectionItems)
            Fail($"{name} exceeds the per-collection item limit.", _limitCode);
        _collectionItems = checked(_collectionItems + count);
        if (_collectionItems > _limits.MaxCollectionItems)
            Fail("Receipt collections exceed the aggregate item limit.", _limitCode);
    }

    public void String(string? value, string name)
    {
        if (value is null)
            return;
        if (value.Length > _limits.MaxStringLength)
            Fail($"{name} exceeds the string-length limit.", _limitCode);
        int utf8Length = Encoding.UTF8.GetByteCount(value);
        if (utf8Length > _limits.MaxStringLength)
            Fail($"{name} exceeds the string-length limit.", _limitCode);
        var encoded = JsonEncodedText.Encode(value, JavaScriptEncoder.Default);
        AddSerializedBytes(encoded.EncodedUtf8Bytes.Length + 2L, name);
    }

    public void AddSerializedBytes(long count, string name)
    {
        if (count < 0 || count > _maximumSerializedBytes - _serializedBytes)
            Fail($"{name} exceeds the aggregate serialized-byte limit.", _limitCode);
        _serializedBytes += count;
    }

    public void Depth(int depth, string name)
    {
        if (depth > _limits.MaxJsonDepth)
            Fail($"{name} exceeds the JSON-depth limit.", _limitCode);
    }

    public static void Bytes(long length, long maximum, string code, string name)
    {
        if (length < 0 || length > maximum)
            throw new DeliveryReceiptValidationException(code, $"{name} exceeds its byte limit.");
    }

    private static void Fail(string message, string code = "receipt_resource_limit") =>
        throw new DeliveryReceiptValidationException(code, message);
}

/// <summary>A MemoryStream that fails before accepting bytes past a configured ceiling.</summary>
internal sealed class DeliveryReceiptBoundedMemoryStream : MemoryStream
{
    private readonly long _maximumBytes;
    private readonly string _limitCode;
    private readonly string _name;

    public DeliveryReceiptBoundedMemoryStream(
        long maximumBytes,
        string limitCode,
        string name)
        : base(capacity: (int)Math.Min(maximumBytes, 16 * 1024))
    {
        _maximumBytes = maximumBytes;
        _limitCode = limitCode;
        _name = name;
    }

    public override void Write(byte[] buffer, int offset, int count)
    {
        EnsureCapacityFor(count);
        base.Write(buffer, offset, count);
    }

    public override void Write(ReadOnlySpan<byte> buffer)
    {
        EnsureCapacityFor(buffer.Length);
        base.Write(buffer);
    }

    public override void WriteByte(byte value)
    {
        EnsureCapacityFor(1);
        base.WriteByte(value);
    }

    public override Task WriteAsync(
        byte[] buffer,
        int offset,
        int count,
        CancellationToken cancellationToken)
    {
        EnsureCapacityFor(count);
        return base.WriteAsync(buffer, offset, count, cancellationToken);
    }

    public override ValueTask WriteAsync(
        ReadOnlyMemory<byte> buffer,
        CancellationToken cancellationToken = default)
    {
        EnsureCapacityFor(buffer.Length);
        return base.WriteAsync(buffer, cancellationToken);
    }

    public override void SetLength(long value)
    {
        if (value > _maximumBytes)
            Fail();
        base.SetLength(value);
    }

    private void EnsureCapacityFor(int count)
    {
        if (count < 0 || Position > _maximumBytes - count)
            Fail();
    }

    private void Fail() => throw new DeliveryReceiptValidationException(
        _limitCode, $"{_name} exceeds its byte limit.");
}
