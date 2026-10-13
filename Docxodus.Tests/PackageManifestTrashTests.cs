#nullable enable
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using Docxodus.Verification;
using Xunit;

namespace Docxodus.Tests;

public class PackageManifestTrashTests
{
    private static byte[] AuthoredPackage(string name, byte[]? payload = null, bool datDefault = false,
        bool explicitOverride = false, bool referenced = false)
    {
        const string rel = "http://schemas.openxmlformats.org/package/2006/relationships";
        var parts = new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = Encoding.UTF8.GetBytes("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/>" + (datDefault ? "<Default Extension=\"dat\" ContentType=\"application/xml\"/>" : "") + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>" + (explicitOverride ? $"<Override PartName=\"/{name}\" ContentType=\"application/octet-stream\"/>" : "") + "</Types>"),
            ["_rels/.rels"] = Encoding.UTF8.GetBytes($"<Relationships xmlns=\"{rel}\"><Relationship Id=\"main\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>" + (referenced ? $"<Relationship Id=\"unused\" Type=\"urn:authored:resource\" Target=\"{name}\"/>" : "") + "</Relationships>"),
            ["word/document.xml"] = Encoding.UTF8.GetBytes("<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p><w:r><w:t>Aspen resource ledger.</w:t></w:r></w:p></w:body></w:document>"),
            [name] = payload ?? Encoding.UTF8.GetBytes("<obsolete>independently authored discarded bytes"),
        };
        if (datDefault) parts["resources/live.dat"] = Encoding.UTF8.GetBytes("<live/>");
        using var stream = new MemoryStream();
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Create, true))
            foreach (var part in parts)
            {
                using var output = zip.CreateEntry(part.Key, CompressionLevel.NoCompression).Open();
                output.Write(part.Value);
            }
        return stream.ToArray();
    }

    [Theory]
    [InlineData("[trash]/0000.dat", false)]
    [InlineData("[trash]/ABCD.dat", false)]
    [InlineData("[trash]/00af.dat", false)]
    [InlineData("[trash]/0000.dat", true)]
    [InlineData("[trash]/ABCD.dat", true)]
    [InlineData("[trash]/00af.dat", true)]
    public void StandardTrashIsOpaqueWithoutAContentType(string name, bool datDefault)
    {
        var bytes = AuthoredPackage(name, datDefault: datDefault);
        var saved = bytes.ToArray();
        var manifest = PackageManifestGenerator.Generate(bytes);
        Assert.True(manifest.IsValid, string.Join("\n", manifest.Findings.Select(f => f.Message)));
        var trash = Assert.Single(manifest.Entries, e => e.Uri == "/" + name);
        Assert.Null(trash.ContentType);
        Assert.False(trash.IsXml);
        Assert.NotNull(trash.RawBytesDigest);
        Assert.Equal(saved, bytes);
        if (datDefault) Assert.Equal("application/xml", Assert.Single(manifest.Entries,
            e => e.Uri == "/resources/live.dat").ContentType);
    }

    [Theory]
    [InlineData("[trash]/unused.bin")]
    [InlineData("[trash]/000.dat")]
    [InlineData("[trash]/00000.dat")]
    [InlineData("[trash]/000g.dat")]
    [InlineData("[trash]/0000.DAT")]
    [InlineData("[Trash]/0000.dat")]
    [InlineData("folder/[trash]/0000.dat")]
    [InlineData("[trash]/../0000.dat")]
    [InlineData("/[trash]/0000.dat")]
    public void OtherBracketedNamesRemainInvalid(string name)
    {
        var manifest = PackageManifestGenerator.Generate(AuthoredPackage(name));
        Assert.False(manifest.IsValid);
        Assert.Contains(manifest.Findings, f => f.Code == "unsafe_entry_path");
    }

    [Fact]
    public void TrashCannotBeDeclaredAsAnOpcPart()
    {
        var manifest = PackageManifestGenerator.Generate(AuthoredPackage("[trash]/0000.dat", explicitOverride: true));
        Assert.False(manifest.IsValid);
        Assert.Contains(manifest.Findings, f => f.Code == "invalid_content_type_part_name");
    }

    [Fact]
    public void TrashCannotBeARelationshipTarget()
    {
        var manifest = PackageManifestGenerator.Generate(AuthoredPackage("[trash]/0000.dat", referenced: true));
        Assert.False(manifest.IsValid);
        Assert.Contains(manifest.Findings, f => f.Code == "invalid_relationship_target");
    }

    [Fact]
    public void TrashBytesStillParticipateInContentAndSemanticDigests()
    {
        var before = PackageManifestGenerator.Generate(AuthoredPackage("[trash]/0000.dat", new byte[] { 1, 2 }));
        var after = PackageManifestGenerator.Generate(AuthoredPackage("[trash]/0000.dat", new byte[] { 1, 3 }));
        Assert.True(before.IsValid && after.IsValid);
        Assert.NotNull(before.OrderedOpcContentDigest);
        Assert.NotNull(before.NormalizedSemanticDigest);
        Assert.NotEqual(before.OrderedOpcContentDigest!.Value, after.OrderedOpcContentDigest?.Value);
        Assert.NotEqual(before.NormalizedSemanticDigest!.Value, after.NormalizedSemanticDigest?.Value);
    }

    [Fact]
    public void TrashChecksumCorruptionStillFails()
    {
        var marker = Encoding.UTF8.GetBytes("IndependentChecksumMarker");
        var bytes = AuthoredPackage("[trash]/0000.dat", marker);
        var index = bytes.AsSpan().IndexOf(marker);
        Assert.True(index > 0);
        bytes[index] ^= 1;
        var manifest = PackageManifestGenerator.Generate(bytes);
        Assert.False(manifest.IsValid);
        Assert.Contains(manifest.Findings, f => f.Code == "crc_mismatch" && f.Location?.EntryUri == "/[trash]/0000.dat");
    }

    [Fact]
    public void TrashEntryAndByteLimitsStillApply()
    {
        var bytes = AuthoredPackage("[trash]/0000.dat", new byte[4096]);
        var sizeLimited = PackageManifestGenerator.Generate(bytes,
            new PackageManifestOptions { MaxEntryUncompressedBytes = 1024 });
        Assert.Contains(sizeLimited.Findings, f => f.Code == "entry_size_limit_exceeded"
            && f.Location?.EntryUri == "/[trash]/0000.dat");
        var countLimited = PackageManifestGenerator.Generate(bytes, new PackageManifestOptions { MaxEntryCount = 3 });
        Assert.Contains(countLimited.Findings, f => f.Code == "entry_count_limit_exceeded");
    }
}
