#nullable enable

// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Buffers.Binary;
using System.IO;
using System.IO.Compression;

namespace Docxodus;

/// <summary>
/// Applies the library's final ZIP policy to an OPC package after its owning Open XML package has
/// finished writing. The package payload is copied entry-for-entry; only ZIP compression and
/// extraction metadata are normalized, and a relationship target the reading funnel encoded
/// (<see cref="PackageOpenRepair"/>) is restored to its original spelling.
/// </summary>
/// <remarks>
/// Word-authored packages commonly mark deflated entries with the ZIP "superfast" hint even when
/// their existing compressed bytes were produced much more efficiently. In update mode,
/// <see cref="ZipArchive"/> maps that hint to <see cref="CompressionLevel.Fastest"/> when an entry
/// is rewritten. A small XML edit can therefore make the compressed package grow far more than its
/// uncompressed content. Rebuilding the completed archive is the first point at which a consistent
/// policy can be applied without mutating a live <c>Package</c> or <c>OpenXmlPackage</c>.
///
/// Entries that were stored because compression did not help remain stored. Package markup uses
/// <see cref="CompressionLevel.Optimal"/> to balance output size and save latency; less frequently
/// written compressible binary parts use <see cref="CompressionLevel.SmallestSize"/> so embedded
/// fonts and similar assets do not expand. Already-compressed media avoids deflate work entirely.
/// </remarks>
internal static class ZipPackageOutputNormalizer
{
    // Unix mode bits for an entry that carries none: what ZipArchive gives a new entry on Linux and
    // macOS, and so on Windows too since the host is fixed (issue #1659).
    internal const int RegularFilePermissions = 0x81A4 << 16; // -rw-r--r-- (0100644)
    private const int DirectoryPermissions = 0x41ED << 16;   // drwxr-xr-x (0040755)

    // Word pins every OPC entry to the ZIP DOS epoch. System.IO.Packaging already emits
    // epoch stamps, but the session's checkpoint clone serializes through ZipArchive,
    // which stamps wall-clock time at 2-second DOS granularity - making the raw package
    // digest of identical logical content time-dependent (issue #521).
    private static readonly DateTimeOffset ZipEpoch = new(1980, 1, 1, 0, 0, 0, TimeSpan.Zero);

    /// <summary>
    /// Returns a normalized copy of <paramref name="packageBytes"/>. Invalid/non-ZIP input passes
    /// through unchanged; callers at this boundary ordinarily supply an already-validated OPC
    /// package.
    /// </summary>
    internal static byte[] Normalize(byte[] packageBytes)
    {
        ArgumentNullException.ThrowIfNull(packageBytes);

        try
        {
            using var sourceStream = new MemoryStream(packageBytes, writable: false);
            using var source = new ZipArchive(sourceStream, ZipArchiveMode.Read);
            using var targetStream = new MemoryStream(packageBytes.Length);

            using (var target = new ZipArchive(targetStream, ZipArchiveMode.Create, leaveOpen: true))
            {
                target.Comment = source.Comment;

                foreach (var sourceEntry in source.Entries)
                {
                    var level = GetCompressionLevel(sourceEntry);
                    var targetEntry = target.CreateEntry(sourceEntry.FullName, level);
                    targetEntry.LastWriteTime = ZipEpoch;
                    targetEntry.Comment = sourceEntry.Comment;
                    targetEntry.ExternalAttributes = NormalizedExternalAttributes(sourceEntry);

                    using var input = sourceEntry.Open();
                    using var output = targetEntry.Open();
                    if (sourceEntry.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
                    {
                        using var rels = new MemoryStream();
                        input.CopyTo(rels);
                        output.Write(PackageOpenRepair.RestoreTargets(rels.ToArray()));
                    }
                    else
                    {
                        input.CopyTo(output);
                    }
                }
            }

            var normalized = targetStream.ToArray();
            StampUnixHost(normalized);
            return normalized;
        }
        catch (InvalidDataException)
        {
            return packageBytes;
        }
    }

    /// <summary>
    /// Records every entry of a finished archive as made on Unix. <see cref="ZipArchive"/> writes the
    /// host it runs on into each central-directory header (0, MS-DOS, on Windows; 3, Unix, elsewhere)
    /// and offers no way to choose it, so the same package differed by a byte per entry between hosts
    /// (issue #1659). Unix, with the permissions <see cref="NormalizedExternalAttributes"/> sets, is
    /// what Linux and macOS already wrote. <paramref name="zip"/> may be only the end of the archive,
    /// from the central directory on, when <paramref name="start"/> gives the offset it begins at.
    /// </summary>
    internal static void StampUnixHost(byte[] zip, long start = 0)
    {
        const byte Unix = 3;
        var bytes = zip.AsSpan();
        // The end-of-central-directory record: the last 22 bytes before an archive comment of at most 64 KB.
        int end = bytes.Length - 22, floor = Math.Max(0, end - 0xFFFF);
        while (end >= floor && BinaryPrimitives.ReadUInt32LittleEndian(bytes[end..]) != 0x06054b50) end--;
        if (end < floor) return;
        long count = BinaryPrimitives.ReadUInt16LittleEndian(bytes[(end + 10)..]);
        long offset = BinaryPrimitives.ReadUInt32LittleEndian(bytes[(end + 16)..]) - start;
        if ((count == 0xFFFF || offset == 0xFFFFFFFF) && end >= 20
            && BinaryPrimitives.ReadUInt32LittleEndian(bytes[(end - 20)..]) == 0x07064b50)
        {
            // Zip64: the locator names the zip64 record, which holds the real count and offset.
            var record = BinaryPrimitives.ReadInt64LittleEndian(bytes[(end - 12)..]) - start;
            if (record < 0 || record + 56 > end) return;
            bytes[(int)record + 13] = Unix;
            count = BinaryPrimitives.ReadInt64LittleEndian(bytes[((int)record + 32)..]);
            offset = BinaryPrimitives.ReadInt64LittleEndian(bytes[((int)record + 48)..]) - start;
        }
        for (long i = 0; i < count && offset >= 0 && offset + 46 <= bytes.Length; i++)
        {
            var header = bytes[(int)offset..];
            if (BinaryPrimitives.ReadUInt32LittleEndian(header) != 0x02014b50) return;
            header[5] = Unix; // "version made by": the specification version, then the host
            offset += 46 + BinaryPrimitives.ReadUInt16LittleEndian(header[28..])
                + BinaryPrimitives.ReadUInt16LittleEndian(header[30..])
                + BinaryPrimitives.ReadUInt16LittleEndian(header[32..]);
        }
    }

    /// <summary>
    /// In-place variant for a caller that owns a <see cref="MemoryStream"/> with no live package or
    /// archive attached. Leaves the stream positioned at zero.
    /// </summary>
    internal static void NormalizeInPlace(MemoryStream stream)
    {
        ArgumentNullException.ThrowIfNull(stream);

        var normalized = Normalize(stream.ToArray());
        stream.SetLength(0);
        stream.Write(normalized, 0, normalized.Length);
        stream.Position = 0;
    }

    private static bool ShouldStore(ZipArchiveEntry entry) =>
        entry.Length == 0 ||
        entry.FullName.EndsWith('/') ||
        (!IsPackageMarkup(entry.FullName) && entry.CompressedLength >= entry.Length);

    private static CompressionLevel GetCompressionLevel(ZipArchiveEntry entry)
    {
        if (ShouldStore(entry))
            return CompressionLevel.NoCompression;

        return IsPackageMarkup(entry.FullName)
            ? CompressionLevel.Optimal
            : CompressionLevel.SmallestSize;
    }

    private static bool IsPackageMarkup(string name) =>
        name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) ||
        name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("[Content_Types].xml", StringComparison.OrdinalIgnoreCase);

    private static int NormalizedExternalAttributes(ZipArchiveEntry entry)
    {
        int permissionBits = (entry.ExternalAttributes >> 16) & 0x1FF;
        if (permissionBits != 0)
            return entry.ExternalAttributes;

        int permissions = entry.FullName.EndsWith('/')
            ? DirectoryPermissions
            : RegularFilePermissions;
        return entry.ExternalAttributes | permissions;
    }
}
