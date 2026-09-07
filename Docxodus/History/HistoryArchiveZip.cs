// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Buffers.Binary;

namespace Docxodus.History;

/// <summary>
/// Bounded ZIP/ZIP64 central-directory preflight BEFORE ZipArchive allocates entry objects.
/// This is a deliberately narrow ZIP profile: one disk, no encryption, stored/deflate entries,
/// no ZIP comments, ASCII manifest/hash names, and bounded extra fields. No filesystem extraction.
/// </summary>
internal static class HistoryArchiveZip
{
    internal static async ValueTask<int> PreflightAsync(Stream input, DocxHistoryArchiveLimits limits, CancellationToken ct)
    {
        HistoryArchiveManifest.Budget(input.Length <= limits.MaxArchiveBytes, "History archive exceeds its byte limit.");
        HistoryArchiveManifest.Require(input.Length >= 22, "Truncated history archive.");
        var end = input.Length - 22;
        var footer = await ReadAtAsync(input, end, 22, ct).ConfigureAwait(false);
        Require(U32(footer, 0) == 0x06054b50 && U16(footer, 20) == 0, "Missing ZIP footer or unsupported ZIP comment.");
        Require(U16(footer, 4) == 0 && U16(footer, 6) == 0, "Split ZIP archives are unsupported.");
        ulong count = U16(footer, 10), centralSize = U32(footer, 12), centralOffset = U32(footer, 16);
        var boundary = (ulong)end;
        if (count == ushort.MaxValue || U16(footer, 8) == ushort.MaxValue
            || centralSize == uint.MaxValue || centralOffset == uint.MaxValue)
        {
            var locator = await ReadAtAsync(input, end - 20, 20, ct).ConfigureAwait(false);
            Require(U32(locator, 0) == 0x07064b50 && U32(locator, 4) == 0 && U32(locator, 16) == 1, "Invalid ZIP64 locator.");
            var position = U64(locator, 8);
            Require(position <= (ulong)Math.Max(0, end - 76), "Invalid ZIP64 footer offset.");
            var wide = await ReadAtAsync(input, (long)position, 56, ct).ConfigureAwait(false);
            Require(U32(wide, 0) == 0x06064b50 && U64(wide, 4) == 44
                && position + 56 == (ulong)(end - 20) && U32(wide, 16) == 0 && U32(wide, 20) == 0,
                "Invalid or unsupported ZIP64 footer.");
            count = U64(wide, 32); centralSize = U64(wide, 40); centralOffset = U64(wide, 48); boundary = position;
            Require(U64(wide, 24) == count, "Split ZIP64 archives are unsupported.");
            Require((U16(footer, 8) == ushort.MaxValue || U16(footer, 8) == count)
                && (U16(footer, 10) == ushort.MaxValue || U16(footer, 10) == count)
                && (U32(footer, 12) == uint.MaxValue || U32(footer, 12) == centralSize)
                && (U32(footer, 16) == uint.MaxValue || U32(footer, 16) == centralOffset), "ZIP footer disagrees with ZIP64 footer.");
        }
        else Require(U16(footer, 8) == count, "Split ZIP archives are unsupported.");
        HistoryArchiveManifest.Budget(count <= (ulong)limits.MaxBlobs + 1 && count <= int.MaxValue, "Too many ZIP entries.");
        Require(count > 0 && centralOffset <= boundary && centralSize == boundary - centralOffset, "Invalid ZIP central-directory range.");
        // Variable fields are bounded per entry before any ZipArchive allocation.
        var cursor = (long)centralOffset;
        for (ulong i = 0; i < count; i++)
        {
            ct.ThrowIfCancellationRequested();
            Require(cursor <= (long)boundary - 46, "Truncated ZIP central-directory entry.");
            var record = await ReadAtAsync(input, cursor, 46, ct).ConfigureAwait(false);
            Require(U32(record, 0) == 0x02014b50, "Invalid ZIP central-directory entry.");
            var flags = U16(record, 8); var method = U16(record, 10);
            Require((flags & ~(8 | 2048 | 6)) == 0 && method is 0 or 8, "Encrypted or unsupported ZIP entry.");
            var nameLength = U16(record, 28); var extraLength = U16(record, 30); var commentLength = U16(record, 32);
            Require(nameLength is > 0 and <= 70 && extraLength <= 128 && commentLength == 0
                && U16(record, 34) == 0, "Unsupported ZIP name, extra field, comment or disk.");
            var next = cursor + 46L + nameLength + extraLength;
            Require(next <= (long)boundary, "ZIP entry exceeds the central directory.");
            var name = await ReadAtAsync(input, cursor + 46, nameLength, ct).ConfigureAwait(false);
            Require(name.All(b => b is >= 32 and < 127), "ZIP entry names must be ASCII.");
            var mode = (U32(record, 38) >> 16) & 0xf000;
            Require(mode is 0 or 0x8000 && (U32(record, 38) & 0x10) == 0, "ZIP directories and links are unsupported.");
            cursor = next;
        }
        Require(cursor == (long)boundary, "Unexpected bytes in ZIP central directory.");
        input.Position = 0;
        return (int)count;
    }

    private static async ValueTask<byte[]> ReadAtAsync(Stream input, long offset, int length, CancellationToken ct)
    {
        Require(offset >= 0 && offset <= input.Length - length, "ZIP metadata points outside the file.");
        input.Position = offset; var bytes = new byte[length];
        await input.ReadExactlyAsync(bytes, ct).ConfigureAwait(false); return bytes;
    }
    private static ushort U16(byte[] bytes, int offset) => BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(offset, 2));
    private static uint U32(byte[] bytes, int offset) => BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(offset, 4));
    private static ulong U64(byte[] bytes, int offset) => BinaryPrimitives.ReadUInt64LittleEndian(bytes.AsSpan(offset, 8));
    private static void Require(bool condition, string message) => HistoryArchiveManifest.Require(condition, message);
}
