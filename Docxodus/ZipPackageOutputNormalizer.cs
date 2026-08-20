#nullable enable

// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
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
    private const int RegularFilePermissions = 0x1A4 << 16; // -rw-r--r-- (0644)
    private const int DirectoryPermissions = 0x1ED << 16;   // drwxr-xr-x (0755)

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
                    input.CopyTo(output);
                }
            }

            return targetStream.ToArray();
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
