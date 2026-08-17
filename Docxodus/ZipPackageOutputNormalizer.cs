#nullable enable

// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Buffers.Binary;
using System.IO;
using System.IO.Compression;

namespace Docxodus;

/// <summary>
/// Applies the library's final ZIP policy to an OPC package after its owning Open XML package has
/// finished writing. The package payload is copied entry-for-entry; only ZIP compression and
/// extraction metadata are normalized.
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
    private const int RegularFilePermissions = unchecked((int)(0x81A4u << 16)); // regular 0644
    private const int DirectoryPermissions = 0x41ED << 16; // directory 0755
    private static readonly DateTimeOffset DeterministicTimestamp =
        new(2000, 1, 1, 0, 0, 0, TimeSpan.Zero);

    /// <summary>
    /// Returns a normalized copy of <paramref name="packageBytes"/>. Invalid/non-ZIP input passes
    /// through unchanged; callers at this boundary ordinarily supply an already-validated OPC
    /// package.
    /// </summary>
    internal static byte[] Normalize(byte[] packageBytes) =>
        Normalize(packageBytes, deterministicContainer: false);

    /// <summary>
    /// Applies the same final ZIP policy while also sorting entries and fixing their container
    /// timestamp. Use this for deterministic generated outputs, never as a semantic-equivalence
    /// substitute for caller-owned packages.
    /// </summary>
    internal static byte[] NormalizeDeterministic(byte[] packageBytes) =>
        Normalize(packageBytes, deterministicContainer: true);

    private static byte[] Normalize(byte[] packageBytes, bool deterministicContainer)
    {
        ArgumentNullException.ThrowIfNull(packageBytes);

        try
        {
            using var sourceStream = new MemoryStream(packageBytes, writable: false);
            using var source = new ZipArchive(sourceStream, ZipArchiveMode.Read);
            using var targetStream = new MemoryStream(packageBytes.Length);

            using (var target = new ZipArchive(targetStream, ZipArchiveMode.Create, leaveOpen: true))
            {
                target.Comment = deterministicContainer ? string.Empty : source.Comment;

                IEnumerable<ZipArchiveEntry> sourceEntries = deterministicContainer
                    ? source.Entries.OrderBy(value => value.FullName, StringComparer.Ordinal)
                    : source.Entries;
                foreach (var sourceEntry in sourceEntries)
                {
                    var level = GetCompressionLevel(sourceEntry, deterministicContainer);
                    var targetEntry = target.CreateEntry(sourceEntry.FullName, level);
                    targetEntry.LastWriteTime = deterministicContainer
                        ? DeterministicTimestamp
                        : sourceEntry.LastWriteTime;
                    targetEntry.Comment = deterministicContainer
                        ? string.Empty
                        : sourceEntry.Comment;
                    targetEntry.ExternalAttributes = NormalizedExternalAttributes(
                        sourceEntry, deterministicContainer);

                    using var input = sourceEntry.Open();
                    using var output = targetEntry.Open();
                    input.CopyTo(output);
                }
            }

            var normalized = targetStream.ToArray();
            if (deterministicContainer)
                NormalizeCentralDirectoryPlatform(normalized);
            return normalized;
        }
        catch (InvalidDataException)
        {
            return packageBytes;
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

    private static CompressionLevel GetCompressionLevel(
        ZipArchiveEntry entry, bool deterministicContainer)
    {
        // In deterministic mode the source's compression ratio is deliberately irrelevant: two
        // semantically identical entries may have arrived stored or deflated by different ZIP
        // writers. Choose only from stable properties of the entry name and payload length.
        if (entry.Length == 0 || entry.FullName.EndsWith('/')
            || (deterministicContainer && IsAlreadyCompressed(entry.FullName))
            || (!deterministicContainer && ShouldStore(entry)))
            return CompressionLevel.NoCompression;

        return IsPackageMarkup(entry.FullName)
            ? CompressionLevel.Optimal
            : CompressionLevel.SmallestSize;
    }

    private static bool IsAlreadyCompressed(string name) =>
        name.EndsWith(".png", StringComparison.OrdinalIgnoreCase) ||
        name.EndsWith(".jpg", StringComparison.OrdinalIgnoreCase) ||
        name.EndsWith(".jpeg", StringComparison.OrdinalIgnoreCase) ||
        name.EndsWith(".gif", StringComparison.OrdinalIgnoreCase) ||
        name.EndsWith(".webp", StringComparison.OrdinalIgnoreCase) ||
        name.EndsWith(".zip", StringComparison.OrdinalIgnoreCase) ||
        name.EndsWith(".gz", StringComparison.OrdinalIgnoreCase) ||
        name.EndsWith(".mp3", StringComparison.OrdinalIgnoreCase) ||
        name.EndsWith(".mp4", StringComparison.OrdinalIgnoreCase) ||
        name.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase);

    private static bool IsPackageMarkup(string name) =>
        name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) ||
        name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("[Content_Types].xml", StringComparison.OrdinalIgnoreCase);

    private static void NormalizeCentralDirectoryPlatform(Span<byte> archive)
    {
        const uint endOfCentralDirectorySignature = 0x06054B50;
        const uint centralDirectoryEntrySignature = 0x02014B50;
        const int endOfCentralDirectoryLength = 22;
        const int maximumZipCommentLength = ushort.MaxValue;
        const int centralDirectoryEntryLength = 46;

        var earliest = Math.Max(0,
            archive.Length - endOfCentralDirectoryLength - maximumZipCommentLength);
        var endOffset = -1;
        for (var offset = archive.Length - endOfCentralDirectoryLength;
             offset >= earliest;
             offset--)
        {
            if (BinaryPrimitives.ReadUInt32LittleEndian(archive.Slice(offset, 4))
                    != endOfCentralDirectorySignature)
                continue;
            var commentLength = BinaryPrimitives.ReadUInt16LittleEndian(
                archive.Slice(offset + 20, 2));
            if (offset + endOfCentralDirectoryLength + commentLength == archive.Length)
            {
                endOffset = offset;
                break;
            }
        }
        if (endOffset < 0)
            throw new InvalidDataException("normalized ZIP has no valid end-of-central-directory record");

        var centralSize = BinaryPrimitives.ReadUInt32LittleEndian(
            archive.Slice(endOffset + 12, 4));
        var centralOffset = BinaryPrimitives.ReadUInt32LittleEndian(
            archive.Slice(endOffset + 16, 4));
        if (centralSize == uint.MaxValue || centralOffset == uint.MaxValue
            || centralOffset > int.MaxValue || centralSize > int.MaxValue
            || (ulong)centralOffset + centralSize > (ulong)endOffset)
            throw new InvalidDataException(
                "deterministic ZIP normalization does not accept ZIP64 central-directory metadata");

        var position = checked((int)centralOffset);
        var centralEnd = checked(position + (int)centralSize);
        while (position < centralEnd)
        {
            if (centralEnd - position < centralDirectoryEntryLength
                || BinaryPrimitives.ReadUInt32LittleEndian(archive.Slice(position, 4))
                    != centralDirectoryEntrySignature)
                throw new InvalidDataException("normalized ZIP central directory is malformed");
            // "Version made by" stores the creator platform in its high byte. ZipArchive.Create
            // uses the current OS, so canonicalize it to the ZIP Unix platform (3) after writing.
            archive[position + 5] = 3;
            var nameLength = BinaryPrimitives.ReadUInt16LittleEndian(
                archive.Slice(position + 28, 2));
            var extraLength = BinaryPrimitives.ReadUInt16LittleEndian(
                archive.Slice(position + 30, 2));
            var commentLength = BinaryPrimitives.ReadUInt16LittleEndian(
                archive.Slice(position + 32, 2));
            position = checked(position + centralDirectoryEntryLength
                + nameLength + extraLength + commentLength);
        }
        if (position != centralEnd)
            throw new InvalidDataException("normalized ZIP central directory length is inconsistent");
    }

    private static int NormalizedExternalAttributes(
        ZipArchiveEntry entry, bool deterministicContainer)
    {
        if (deterministicContainer)
        {
            // Do not let the source host or checkout umask leak into reproducible packages.
            return entry.FullName.EndsWith('/')
                ? DirectoryPermissions | 0x10
                : RegularFilePermissions;
        }
        if (OperatingSystem.IsWindows())
            return entry.ExternalAttributes;

        int permissionBits = (entry.ExternalAttributes >> 16) & 0x1FF;
        if (permissionBits != 0)
            return entry.ExternalAttributes;

        int permissions = entry.FullName.EndsWith('/')
            ? DirectoryPermissions
            : RegularFilePermissions;
        return entry.ExternalAttributes | permissions;
    }
}
