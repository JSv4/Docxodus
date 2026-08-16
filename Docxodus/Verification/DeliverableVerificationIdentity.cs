// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Buffers.Binary;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;

namespace Docxodus.Verification;

internal static class DeliverableVerificationIdentity
{
    internal static VerificationDigest Digest(ReadOnlySpan<byte> bytes)
    {
        var hash = SHA256.HashData(bytes);
        return new VerificationDigest
        {
            Algorithm = "SHA-256",
            Value = Convert.ToHexString(hash).ToLowerInvariant(),
        };
    }

    internal static string Token(string domain, params string?[] values)
    {
        using var hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        Append(hash, domain);
        foreach (var value in values)
            Append(hash, value);
        return Convert.ToHexString(hash.GetHashAndReset()).ToLowerInvariant();
    }

    internal static string SemanticChangeFingerprint(SemanticChange change) => Token(
        "docxodus.deliverable.semantic-change.v1",
        ((int)change.Operation).ToString(CultureInfo.InvariantCulture),
        ((int)change.Family).ToString(CultureInfo.InvariantCulture),
        change.PartUri,
        change.Path,
        change.LeftAnchor,
        change.RightAnchor,
        change.LeftScope,
        change.RightScope,
        change.MoveId,
        SemanticValueKey(change.Before),
        SemanticValueKey(change.After));

    internal static string LocationKey(ChangeLocation? location) => string.Join("\u001f", new[]
    {
        location?.EntryUri ?? string.Empty,
        location?.OwnerUri ?? string.Empty,
        location?.RelationshipId ?? string.Empty,
        location?.TargetUri ?? string.Empty,
        location?.PropertyPath ?? string.Empty,
    });

    internal static bool DigestEquals(VerificationDigest? left, VerificationDigest? right) =>
        left is null && right is null
        || left is not null && right is not null
        && string.Equals(left.Algorithm, right.Algorithm, StringComparison.OrdinalIgnoreCase)
        && string.Equals(left.Value, right.Value, StringComparison.OrdinalIgnoreCase);

    internal static string SanitizeCode(string value)
    {
        if (string.IsNullOrWhiteSpace(value)) return "unknown";
        var builder = new StringBuilder(value.Length);
        bool separator = false;
        foreach (var character in value)
        {
            if (char.IsAsciiLetterOrDigit(character))
            {
                builder.Append(char.ToLowerInvariant(character));
                separator = false;
            }
            else if (!separator && builder.Length > 0)
            {
                builder.Append('_');
                separator = true;
            }
        }
        while (builder.Length > 0 && builder[^1] == '_') builder.Length--;
        return builder.Length == 0 ? "unknown" : builder.ToString();
    }

    private static string SemanticValueKey(SemanticValue value)
    {
        var builder = new StringBuilder();
        AppendValue(builder, value);
        return builder.ToString();
    }

    private static void AppendValue(StringBuilder builder, SemanticValue value)
    {
        builder.Append(((int)value.Kind).ToString(CultureInfo.InvariantCulture)).Append(':');
        switch (value.Kind)
        {
            case SemanticValueKind.Absent:
                return;
            case SemanticValueKind.String:
                AppendString(builder, value.StringValue);
                return;
            case SemanticValueKind.Boolean:
                builder.Append(value.BooleanValue == true ? '1' : '0');
                return;
            case SemanticValueKind.Integer:
                builder.Append(value.IntegerValue?.ToString(CultureInfo.InvariantCulture));
                return;
            case SemanticValueKind.Digest:
                AppendString(builder, value.DigestAlgorithm);
                AppendString(builder, value.DigestProfile);
                AppendString(builder, value.DigestValue);
                return;
            case SemanticValueKind.Object:
                foreach (var property in value.Properties)
                {
                    AppendString(builder, property.Name);
                    AppendValue(builder, property.Value);
                }
                return;
            case SemanticValueKind.Array:
                foreach (var item in value.Items)
                    AppendValue(builder, item);
                return;
            default:
                throw new ArgumentOutOfRangeException(nameof(value), value.Kind, null);
        }
    }

    private static void AppendString(StringBuilder builder, string? value)
    {
        if (value is null)
        {
            builder.Append("-1:");
            return;
        }
        builder.Append(value.Length.ToString(CultureInfo.InvariantCulture)).Append(':').Append(value);
    }

    private static void Append(IncrementalHash hash, string? value)
    {
        if (value is null)
        {
            Span<byte> missing = stackalloc byte[4];
            BinaryPrimitives.WriteInt32LittleEndian(missing, -1);
            hash.AppendData(missing);
            return;
        }

        var bytes = Encoding.UTF8.GetBytes(value);
        Span<byte> length = stackalloc byte[4];
        BinaryPrimitives.WriteInt32LittleEndian(length, bytes.Length);
        hash.AppendData(length);
        hash.AppendData(bytes);
    }
}
