// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Buffers.Binary;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Docxodus.Verification;
using Xunit;

namespace OxPt;

public class PackageManifestOleAndContentTypeTests
{
    private const uint DifSector = 0xfffffffc;
    private const uint FatSector = 0xfffffffd;
    private const uint EndOfChain = 0xfffffffe;
    private const uint FreeSector = 0xffffffff;
    private static readonly DateTimeOffset StableTimestamp =
        new(2020, 1, 1, 0, 0, 0, TimeSpan.Zero);

    [Fact]
    public void OleEncryptionDetection_RequiresBothValidatedDirectoryStreams()
    {
        var encrypted = PackageManifestGenerator.Generate(
            BuildCompoundFile("EncryptedPackage", "EncryptionInfo"));
        var caseVariant = PackageManifestGenerator.Generate(
            BuildCompoundFile("encryptedpackage", "ENCRYPTIONINFO"));
        var oneNamedStream = PackageManifestGenerator.Generate(
            BuildCompoundFile("EncryptedPackage", "OtherStream"));
        var generic = PackageManifestGenerator.Generate(
            BuildCompoundFile("Workbook", "SummaryInformation"));

        Assert.Equal("ole-encrypted", encrypted.PackageKind);
        Assert.Contains(encrypted.Findings,
            finding => finding.Code == "unsupported_ole_encryption");
        Assert.Equal("ole-encrypted", caseVariant.PackageKind);
        Assert.Equal("ole", oneNamedStream.PackageKind);
        Assert.Equal("ole", generic.PackageKind);
        Assert.All(new[] { oneNamedStream, generic }, manifest =>
            Assert.Contains(manifest.Findings,
                finding => finding.Code == "unsupported_compound_file"));
    }

    [Fact]
    public void OleEncryptionDetection_ValidatesMiniFatAndRootMiniStreamChains()
    {
        var validMiniStreams = BuildMiniStreamCompoundFile();
        var withoutMiniFat = BuildMiniStreamCompoundFile();
        WriteUInt32(withoutMiniFat, 60, EndOfChain);
        WriteUInt32(withoutMiniFat, 64, 0);
        var corruptMiniFat = BuildMiniStreamCompoundFile();
        const int miniFatOffset = 3 * 512;
        WriteUInt32(corruptMiniFat, miniFatOffset + sizeof(uint), FreeSector);
        var corruptRootMiniStream = BuildMiniStreamCompoundFile();
        const int fatOffset = 512;
        WriteUInt32(corruptRootMiniStream,
            fatOffset + 3 * sizeof(uint), FreeSector);

        var encrypted = PackageManifestGenerator.Generate(validMiniStreams);
        Assert.Equal("ole-encrypted", encrypted.PackageKind);
        Assert.Contains(encrypted.Findings,
            finding => finding.Code == "unsupported_ole_encryption");

        var malformedMiniStreams = new[]
        {
            withoutMiniFat,
            corruptMiniFat,
            corruptRootMiniStream,
        };
        foreach (var bytes in malformedMiniStreams)
        {
            var manifest = PackageManifestGenerator.Generate(bytes);
            Assert.Equal("ole", manifest.PackageKind);
            Assert.DoesNotContain(manifest.Findings,
                finding => finding.Code == "unsupported_ole_encryption");
            Assert.Contains(manifest.Findings,
                finding => finding.Code == "unsupported_compound_file");
        }
    }

    [Fact]
    public void OleEncryptionDetection_DoesNotScanPayloadOrCorruptHeaderSubstrings()
    {
        var namesInPayload = BuildCompoundFile("Workbook", "SummaryInformation");
        Encoding.Unicode.GetBytes("EncryptedPackage").CopyTo(namesInPayload, 1536);
        Encoding.Unicode.GetBytes("EncryptionInfo").CopyTo(namesInPayload, 1600);

        var corrupt = new byte[512];
        new byte[] { 0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1 }
            .CopyTo(corrupt, 0);
        Encoding.Unicode.GetBytes("EncryptedPackage").CopyTo(corrupt, 64);
        Encoding.Unicode.GetBytes("EncryptionInfo").CopyTo(corrupt, 128);

        var orphanDirectoryEntries = BuildCompoundFile("Workbook", "OtherStream");
        const int directoryOffset = 1024;
        WriteDirectoryEntry(orphanDirectoryEntries, directoryOffset + 128, "Workbook", 2, 1,
            FreeSector, FreeSector, FreeSector, 2, 4096);
        WriteDirectoryEntry(orphanDirectoryEntries, directoryOffset + 256,
            "EncryptedPackage", 2, 1,
            FreeSector, FreeSector, FreeSector, 2, 4096);
        WriteDirectoryEntry(orphanDirectoryEntries, directoryOffset + 384,
            "EncryptionInfo", 2, 1,
            FreeSector, FreeSector, FreeSector, 10, 4096);

        var overlappingEncryptionStreams = BuildCompoundFile(
            "EncryptedPackage", "EncryptionInfo");
        WriteDirectoryEntry(overlappingEncryptionStreams, directoryOffset + 256,
            "EncryptionInfo", 2, 0,
            FreeSector, FreeSector, FreeSector, 2, 4096);

        foreach (var bytes in new[]
                 { namesInPayload, orphanDirectoryEntries, overlappingEncryptionStreams, corrupt })
        {
            var manifest = PackageManifestGenerator.Generate(bytes);
            Assert.Equal("ole", manifest.PackageKind);
            Assert.DoesNotContain(manifest.Findings,
                finding => finding.Code == "unsupported_ole_encryption");
            Assert.Contains(manifest.Findings,
                finding => finding.Code == "unsupported_compound_file");
        }
    }

    [Theory]
    [InlineData("application")]
    [InlineData("application/")]
    [InlineData("/xml")]
    [InlineData("application/xml ")]
    [InlineData(" application/xml")]
    [InlineData("application\u00a0/xml")]
    [InlineData("application /xml")]
    [InlineData("application/xml;")]
    [InlineData("application/xml; charset")]
    [InlineData("application/xml; charset =utf-8")]
    [InlineData("application/xml; charset= utf-8")]
    [InlineData("application/xml; charset=\"unterminated")]
    [InlineData("application/(xml)")]
    public void ContentTypeMediaType_MalformedSyntaxIsStructuredAndNotResolved(string contentType)
    {
        var bytes = BuildOpc(contentType);

        var manifest = PackageManifestGenerator.Generate(bytes);
        var repeated = PackageManifestGenerator.Generate(bytes);

        var finding = Assert.Single(manifest.Findings,
            item => item.Code == "malformed_content_type");
        Assert.Equal("/[Content_Types].xml", finding.Location?.EntryUri);
        Assert.Equal("default:xml", finding.Location?.PropertyPath);
        Assert.Equal(contentType, Assert.Single(manifest.ContentTypes).ContentType);
        Assert.Null(manifest.Entries.Single(entry => entry.Uri == "/word/document.xml").ContentType);
        Assert.Equal(manifest.ToJson(), repeated.ToJson());
    }

    [Theory]
    [InlineData("application/xml")]
    [InlineData("application/vnd.example+xml")]
    [InlineData("application/vnd.example+xml;charset=utf-8")]
    [InlineData("application/vnd.example+xml ; charset=utf-8")]
    [InlineData("application/vnd.example+xml;charset=\"utf-8\"")]
    public void ContentTypeMediaType_ValidSyntaxStillResolves(string contentType)
    {
        var manifest = PackageManifestGenerator.Generate(BuildOpc(contentType));

        Assert.DoesNotContain(manifest.Findings,
            finding => finding.Code == "malformed_content_type");
        Assert.Equal(contentType,
            manifest.Entries.Single(entry => entry.Uri == "/word/document.xml").ContentType);
    }

    private static byte[] BuildOpc(string contentType)
    {
        XNamespace contentTypes =
            "http://schemas.openxmlformats.org/package/2006/content-types";
        var contentTypesXml = new XDocument(
            new XElement(contentTypes + "Types",
                new XElement(contentTypes + "Default",
                    new XAttribute("Extension", "xml"),
                    new XAttribute("ContentType", contentType))));

        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
        {
            AddZipEntry(archive, "[Content_Types].xml",
                Encoding.UTF8.GetBytes(contentTypesXml.ToString(SaveOptions.DisableFormatting)));
            AddZipEntry(archive, "word/document.xml",
                Encoding.UTF8.GetBytes("<document xmlns=\"urn:test\"/>"));
        }
        return output.ToArray();
    }

    private static void AddZipEntry(ZipArchive archive, string name, byte[] payload)
    {
        var entry = archive.CreateEntry(name, CompressionLevel.NoCompression);
        entry.LastWriteTime = StableTimestamp;
        using var stream = entry.Open();
        stream.Write(payload);
    }

    private static byte[] BuildCompoundFile(string firstStreamName, string secondStreamName)
    {
        const int sectorSize = 512;
        const int sectorCount = 18;
        var bytes = new byte[(sectorCount + 1) * sectorSize];
        new byte[] { 0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1 }
            .CopyTo(bytes, 0);

        WriteUInt16(bytes, 24, 0x003e);
        WriteUInt16(bytes, 26, 3);
        WriteUInt16(bytes, 28, 0xfffe);
        WriteUInt16(bytes, 30, 9);
        WriteUInt16(bytes, 32, 6);
        WriteUInt32(bytes, 40, 0);
        WriteUInt32(bytes, 44, 1);
        WriteUInt32(bytes, 48, 1);
        WriteUInt32(bytes, 56, 4096);
        WriteUInt32(bytes, 60, EndOfChain);
        WriteUInt32(bytes, 64, 0);
        WriteUInt32(bytes, 68, EndOfChain);
        WriteUInt32(bytes, 72, 0);
        for (var index = 0; index < 109; index++)
            WriteUInt32(bytes, 76 + index * sizeof(uint), FreeSector);
        WriteUInt32(bytes, 76, 0);

        var fatOffset = sectorSize;
        for (var index = 0; index < sectorSize / sizeof(uint); index++)
            WriteUInt32(bytes, fatOffset + index * sizeof(uint), FreeSector);
        WriteUInt32(bytes, fatOffset, FatSector);
        WriteUInt32(bytes, fatOffset + sizeof(uint), EndOfChain);
        WriteSectorChain(bytes, fatOffset, 2, 9);
        WriteSectorChain(bytes, fatOffset, 10, 17);

        var directoryOffset = sectorSize * 2;
        WriteDirectoryEntry(bytes, directoryOffset, "Root Entry", 5, 1,
            FreeSector, FreeSector, 1, EndOfChain, 0);
        var secondComesBeforeFirst = CompareCompoundFileNames(
            secondStreamName, firstStreamName) < 0;
        WriteDirectoryEntry(bytes, directoryOffset + 128, firstStreamName, 2, 1,
            secondComesBeforeFirst ? 2u : FreeSector,
            secondComesBeforeFirst ? FreeSector : 2u,
            FreeSector, 2, 4096);
        WriteDirectoryEntry(bytes, directoryOffset + 256, secondStreamName, 2, 0,
            FreeSector, FreeSector, FreeSector, 10, 4096);
        return bytes;
    }

    private static byte[] BuildMiniStreamCompoundFile()
    {
        const int sectorSize = 512;
        const int sectorCount = 4;
        var bytes = new byte[(sectorCount + 1) * sectorSize];
        new byte[] { 0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1 }
            .CopyTo(bytes, 0);

        WriteUInt16(bytes, 24, 0x003e);
        WriteUInt16(bytes, 26, 3);
        WriteUInt16(bytes, 28, 0xfffe);
        WriteUInt16(bytes, 30, 9);
        WriteUInt16(bytes, 32, 6);
        WriteUInt32(bytes, 40, 0);
        WriteUInt32(bytes, 44, 1);
        WriteUInt32(bytes, 48, 1);
        WriteUInt32(bytes, 56, 4096);
        WriteUInt32(bytes, 60, 2);
        WriteUInt32(bytes, 64, 1);
        WriteUInt32(bytes, 68, EndOfChain);
        WriteUInt32(bytes, 72, 0);
        for (var index = 0; index < 109; index++)
            WriteUInt32(bytes, 76 + index * sizeof(uint), FreeSector);
        WriteUInt32(bytes, 76, 0);

        var fatOffset = sectorSize;
        for (var index = 0; index < sectorSize / sizeof(uint); index++)
            WriteUInt32(bytes, fatOffset + index * sizeof(uint), FreeSector);
        WriteUInt32(bytes, fatOffset, FatSector);
        WriteUInt32(bytes, fatOffset + sizeof(uint), EndOfChain);
        WriteUInt32(bytes, fatOffset + 2 * sizeof(uint), EndOfChain);
        WriteUInt32(bytes, fatOffset + 3 * sizeof(uint), EndOfChain);

        var directoryOffset = 2 * sectorSize;
        WriteDirectoryEntry(bytes, directoryOffset, "Root Entry", 5, 1,
            FreeSector, FreeSector, 1, 3, 128);
        WriteDirectoryEntry(bytes, directoryOffset + 128, "EncryptedPackage", 2, 1,
            2, FreeSector, FreeSector, 0, 8);
        WriteDirectoryEntry(bytes, directoryOffset + 256, "EncryptionInfo", 2, 0,
            FreeSector, FreeSector, FreeSector, 1, 8);

        var miniFatOffset = 3 * sectorSize;
        for (var index = 0; index < sectorSize / sizeof(uint); index++)
            WriteUInt32(bytes, miniFatOffset + index * sizeof(uint), FreeSector);
        WriteUInt32(bytes, miniFatOffset, EndOfChain);
        WriteUInt32(bytes, miniFatOffset + sizeof(uint), EndOfChain);
        return bytes;
    }

    private static int CompareCompoundFileNames(string left, string right)
    {
        var byLength = left.Length.CompareTo(right.Length);
        return byLength != 0
            ? byLength
            : StringComparer.OrdinalIgnoreCase.Compare(left, right);
    }

    private static void WriteSectorChain(byte[] bytes, int fatOffset, int first, int last)
    {
        for (var sector = first; sector <= last; sector++)
        {
            WriteUInt32(bytes, fatOffset + sector * sizeof(uint),
                sector == last ? EndOfChain : (uint)(sector + 1));
        }
    }

    private static void WriteDirectoryEntry(
        byte[] bytes,
        int offset,
        string name,
        byte objectType,
        byte color,
        uint leftSibling,
        uint rightSibling,
        uint child,
        uint startingSector,
        ulong streamSize)
    {
        var encodedName = Encoding.Unicode.GetBytes(name);
        encodedName.CopyTo(bytes, offset);
        WriteUInt16(bytes, offset + 64, (ushort)(encodedName.Length + sizeof(ushort)));
        bytes[offset + 66] = objectType;
        bytes[offset + 67] = color;
        WriteUInt32(bytes, offset + 68, leftSibling);
        WriteUInt32(bytes, offset + 72, rightSibling);
        WriteUInt32(bytes, offset + 76, child);
        WriteUInt32(bytes, offset + 116, startingSector);
        BinaryPrimitives.WriteUInt64LittleEndian(bytes.AsSpan(offset + 120), streamSize);
    }

    private static void WriteUInt16(byte[] bytes, int offset, ushort value) =>
        BinaryPrimitives.WriteUInt16LittleEndian(bytes.AsSpan(offset), value);

    private static void WriteUInt32(byte[] bytes, int offset, uint value) =>
        BinaryPrimitives.WriteUInt32LittleEndian(bytes.AsSpan(offset), value);
}
