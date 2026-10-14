#nullable enable
using System.IO.Compression;
using System.Text;
using Docxodus.Verification;
using Xunit;

namespace Docxodus.Tests;

public class PackageManifestThumbnailTests
{
    // Original 2x2 RGB tile encoded from literal pixel values; no source image or DOCX.
    private static readonly byte[] Tile = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAE0lEQVR4nGNwSes4c/cdAxADWQA0QAeTbjIQcQAAAABJRU5ErkJggg==");

    private static byte[] Package(string owner, string target, bool present, bool relationship = true,
        byte[]? payload = null)
    {
        const string rel = "http://schemas.openxmlformats.org/package/2006/relationships";
        const string office = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var thumbnail = relationship
            ? $"<Relationship Id=\"preview\" Type=\"{rel}/metadata/thumbnail\" Target=\"{target}\"/>" : "";
        var parts = new Dictionary<string, byte[]>
        {
            ["[Content_Types].xml"] = Encoding.UTF8.GetBytes("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"png\" ContentType=\"image/png\"/><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/></Types>"),
            ["_rels/.rels"] = Encoding.UTF8.GetBytes($"<Relationships xmlns=\"{rel}\"><Relationship Id=\"main\" Type=\"{office}/officeDocument\" Target=\"word/document.xml\"/>" + (owner == "/" ? thumbnail : "") + "</Relationships>"),
            ["word/document.xml"] = Encoding.UTF8.GetBytes("<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p><w:r><w:t>Maple records remain complete.</w:t></w:r></w:p></w:body></w:document>"),
        };
        if (owner != "/") parts["word/_rels/document.xml.rels"] = Encoding.UTF8.GetBytes(
            $"<Relationships xmlns=\"{rel}\">{thumbnail}</Relationships>");
        if (present) parts["docProps/preview.png"] = payload ?? Tile;
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
    [InlineData("/", "docProps/preview.png", "/_rels/.rels")]
    [InlineData("/", "/docProps/preview.png", "/_rels/.rels")]
    [InlineData("/word/document.xml", "../docProps/preview.png", "/word/_rels/document.xml.rels")]
    public void MissingThumbnailRetainsItsRelationshipAndPreciseError(string owner, string target, string entry)
    {
        var bytes = Package(owner, target, false);
        var saved = bytes.ToArray();
        var manifest = PackageManifestGenerator.Generate(bytes);
        Assert.False(manifest.IsValid);
        var finding = Assert.Single(manifest.Findings);
        Assert.Equal("missing_target", finding.Code);
        Assert.Equal(VerificationFindingSeverity.Error, finding.Severity);
        Assert.Equal(entry, finding.Location?.EntryUri);
        Assert.Equal(owner, finding.Location?.OwnerUri);
        Assert.Equal("preview", finding.Location?.RelationshipId);
        Assert.Equal("/docProps/preview.png", finding.Location?.TargetUri);
        var thumbnail = Assert.Single(manifest.Relationships, r => r.Id == "preview");
        Assert.Equal("Internal", thumbnail.TargetMode);
        Assert.Equal("/docProps/preview.png", thumbnail.ResolvedTargetUri);
        Assert.False(thumbnail.IsTargetPresent);
        Assert.Equal(saved, bytes);
    }

    [Theory]
    [InlineData("/", "docProps/preview.png")]
    [InlineData("/", "/docProps/preview.png")]
    [InlineData("/word/document.xml", "../docProps/preview.png")]
    public void PresentThumbnailIsValidOpaqueImageData(string owner, string target)
    {
        var manifest = PackageManifestGenerator.Generate(Package(owner, target, true));
        Assert.True(manifest.IsValid, string.Join("\n", manifest.Findings.Select(f => f.Message)));
        Assert.True(Assert.Single(manifest.Relationships, r => r.Id == "preview").IsTargetPresent);
        var image = Assert.Single(manifest.Entries, e => e.Uri == "/docProps/preview.png");
        Assert.Equal("image/png", image.ContentType);
        Assert.False(image.IsXml);
        Assert.Equal(Tile.Length, image.Size);
        Assert.NotNull(image.RawBytesDigest);
    }

    [Fact]
    public void AnOmittedThumbnailDoesNotRequireAnImagePart()
    {
        var manifest = PackageManifestGenerator.Generate(Package("/", "docProps/preview.png", false, false));
        Assert.True(manifest.IsValid);
        Assert.DoesNotContain(manifest.Relationships, r => r.Id == "preview");
    }

    [Fact]
    public void ThumbnailChecksumAndSizeLimitsStillApply()
    {
        var bytes = Package("/", "docProps/preview.png", true);
        var index = bytes.AsSpan().IndexOf(Tile);
        Assert.True(index > 0);
        bytes[index + 20] ^= 1;
        var corrupt = PackageManifestGenerator.Generate(bytes);
        Assert.False(corrupt.IsValid);
        Assert.Contains(corrupt.Findings, f => f.Code == "crc_mismatch"
            && f.Location?.EntryUri == "/docProps/preview.png");
        var limited = PackageManifestGenerator.Generate(
            Package("/", "docProps/preview.png", true, payload: new byte[4096]),
            new PackageManifestOptions { MaxEntryUncompressedBytes = 1024 });
        Assert.False(limited.IsValid);
        Assert.Contains(limited.Findings, f => f.Code == "entry_size_limit_exceeded"
            && f.Location?.EntryUri == "/docProps/preview.png");
    }
}
