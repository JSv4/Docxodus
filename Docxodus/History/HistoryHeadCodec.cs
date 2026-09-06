// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Docxodus.Verification;

namespace Docxodus.History;

internal static class HistoryHeadCodec
{
    internal const int MaxBytes = 1024;

    internal static string Key(string documentId)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(documentId);
        if (documentId.Length > 1024) throw new ArgumentException("Document ID exceeds the limit.", nameof(documentId));
        // Reject malformed UTF-16 instead of mapping distinct invalid IDs to replacement bytes.
        var bytes = new UTF8Encoding(false, true).GetBytes(documentId);
        return Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
    }

    internal static HistoryHead Next(HistoryHead? expected, HistoryBlobReference state)
    {
        HistoryBlobIO.Validate(state, int.MaxValue);
        if (expected is not null) Validate(expected);
        return new HistoryHead(checked((expected?.Revision ?? 0) + 1), state);
    }

    internal static void Validate(HistoryHead head)
    {
        ArgumentNullException.ThrowIfNull(head);
        HistoryBlobIO.Validate(head.State, int.MaxValue);
        if (head.Revision <= 0) throw new ArgumentOutOfRangeException(nameof(head));
    }

    internal static byte[] Encode(HistoryHead head)
    {
        using var buffer = new MemoryStream();
        using (var writer = new Utf8JsonWriter(buffer))
        {
            writer.WriteStartObject();
            writer.WriteString("digest", head.State.Digest.Value);
            writer.WriteNumber("length", head.State.Length);
            writer.WriteNumber("revision", head.Revision);
            writer.WriteNumber("schemaVersion", 1);
            writer.WriteEndObject();
        }
        return buffer.ToArray();
    }

    internal static HistoryHead Decode(byte[] bytes)
    {
        try
        {
            if (bytes.Length > MaxBytes) throw new InvalidDataException("History head exceeds the byte limit.");
            using var document = JsonDocument.Parse(bytes, new JsonDocumentOptions { MaxDepth = 2 });
            var root = document.RootElement;
            var properties = new HashSet<string>(StringComparer.Ordinal);
            foreach (var property in root.EnumerateObject())
                if (property.Name is not ("digest" or "length" or "revision" or "schemaVersion") || !properties.Add(property.Name))
                    throw new InvalidDataException("Unknown or duplicate history head property.");
            if (properties.Count != 4 || root.GetProperty("schemaVersion").GetInt32() != 1)
                throw new InvalidDataException("Missing or unsupported history head schema.");
            var head = new HistoryHead(root.GetProperty("revision").GetInt64(),
                new HistoryBlobReference(new VerificationDigest
                {
                    Algorithm = "SHA-256", Value = root.GetProperty("digest").GetString()!,
                }, root.GetProperty("length").GetInt32()));
            HistoryBlobIO.Validate(head.State, int.MaxValue);
            if (head.Revision <= 0) throw new InvalidDataException("Invalid publication revision.");
            return head;
        }
        catch (Exception error) when (error is JsonException or InvalidOperationException or FormatException or ArgumentException)
        {
            throw new InvalidDataException("Malformed history head.", error);
        }
    }
}
