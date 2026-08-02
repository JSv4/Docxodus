#nullable enable

using System;
using System.IO;
using System.IO.Compression;

namespace Docxodus;

/// <summary>
/// Mitigates a <see cref="System.IO.Packaging"/> quirk (issue #302): re-saving an OPC package that
/// was opened from existing bytes stamps the zip central directory's "version made by" host as Unix
/// on a non-Windows OS, but leaves <c>ExternalAttributes</c> (the Unix permission bits) at 0 for
/// entries the package writer didn't explicitly set. A standard Unix <c>unzip</c> takes a
/// Unix-hosted, zero-permission entry literally and extracts it as mode <c>000</c> (unreadable, even
/// to the owner) — Word, LibreOffice, and Python's <c>zipfile</c> module don't consult these bits at
/// all, so the corruption is invisible until something shells out to <c>unzip</c>. See
/// docs/ooxml_corner_cases.md for the full repro.
/// </summary>
internal static class ZipUnixPermissionFixer
{
    private const int RegularFilePermissions = 0x1A4 << 16; // -rw-r--r-- (0644)
    private const int DirectoryPermissions = 0x1ED << 16;   // drwxr-xr-x (0755)

    /// <summary>
    /// Returns a copy of <paramref name="documentBytes"/> with sane Unix permission bits assigned to
    /// any zip entry currently at 0. No-ops (returns the same array) on Windows, where extraction
    /// tools don't consult these bits and the host byte a fresh save gets stamped with is DOS, not
    /// Unix. Safe to call on any OPC package (docx/xlsx/pptx) — non-zip or already-correct inputs
    /// pass through unchanged.
    /// </summary>
    internal static byte[] Fix(byte[] documentBytes)
    {
        if (OperatingSystem.IsWindows())
            return documentBytes;

        using var ms = new MemoryStream();
        ms.Write(documentBytes, 0, documentBytes.Length);
        FixInPlace(ms);
        return ms.ToArray();
    }

    /// <summary>
    /// In-place variant for callers that already own a <see cref="MemoryStream"/> with no other live
    /// reader/writer attached to it (e.g. a package that was just disposed). Do NOT call this on a
    /// stream that a long-lived <see cref="System.IO.Packaging.Package"/> or
    /// <c>WordprocessingDocument</c> may still write to later — rewriting the physical zip layout out
    /// from under a still-open package's own archive handle risks corrupting its next save. Leaves
    /// the stream positioned at 0 on return either way.
    /// </summary>
    internal static void FixInPlace(MemoryStream stream)
    {
        stream.Position = 0;
        if (OperatingSystem.IsWindows())
            return;

        using (var zip = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            foreach (var entry in zip.Entries)
            {
                if (entry.ExternalAttributes != 0)
                    continue;
                bool isDirectory = entry.FullName.EndsWith('/');
                entry.ExternalAttributes = isDirectory ? DirectoryPermissions : RegularFilePermissions;
            }
        }
        stream.Position = 0;
    }
}
