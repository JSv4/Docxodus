#nullable enable
using System.IO.Compression;
using System.Text;
using Docxodus.Verification;
using Xunit;

namespace Docxodus.Tests;

public class PackageManifestCustomPropertiesTests
{
    private static string Metadata(string kind)
    {
        var value = kind switch
        {
            "number" => "<vt:i4>37</vt:i4>",
            "boolean" => "<vt:bool>true</vt:bool>",
            "entity" => "<vt:lpwstr>&missing;</vt:lpwstr>",
            "large" => $"<vt:lpwstr>{new string('x', 4096)}</vt:lpwstr>",
            _ => "<vt:lpwstr>CedarMetadataOnly &amp; &lt;kept&gt;.</vt:lpwstr>",
        };
        var text = "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/custom-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\"><property fmtid=\"{D5CDD505-2E9C-101B-9397-08002B2CF9AE}\" pid=\"2\" name=\"CedarRecord\">" + value + "</property></Properties>";
        return kind switch
        {
            "mismatched" => text.Replace("</property>", "</wrong>"),
            "truncated" => text.Replace("</Properties>", ""),
            _ => text,
        };
    }

    private static byte[] Package(string kind)
    {
        const string rel = "http://schemas.openxmlformats.org/package/2006/relationships";
        const string office = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var present = kind != "none";
        var parts = new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>" + (present ? "<Override PartName=\"/docProps/custom.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.custom-properties+xml\"/>" : "") + "</Types>",
            ["_rels/.rels"] = $"<Relationships xmlns=\"{rel}\"><Relationship Id=\"main\" Type=\"{office}/officeDocument\" Target=\"word/document.xml\"/>" + (present ? $"<Relationship Id=\"custom\" Type=\"{office}/custom-properties\" Target=\"docProps/custom.xml\"/>" : "") + "</Relationships>",
            ["word/document.xml"] = "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p><w:r><w:t>Cedar body records stay visible.</w:t></w:r></w:p></w:body></w:document>",
        };
        if (present) parts["docProps/custom.xml"] = Metadata(kind);
        using var stream = new MemoryStream();
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Create, true))
            foreach (var part in parts)
            {
                using var output = zip.CreateEntry(part.Key, CompressionLevel.NoCompression).Open();
                output.Write(Encoding.UTF8.GetBytes(part.Value));
            }
        return stream.ToArray();
    }

    [Theory]
    [InlineData("mismatched")]
    [InlineData("truncated")]
    [InlineData("entity")]
    public void MalformedCustomPropertiesAreNamedAndNeverTreatedAsNormalizedXml(string kind)
    {
        var bytes = Package(kind);
        var saved = bytes.ToArray();
        var manifest = PackageManifestGenerator.Generate(bytes);
        Assert.False(manifest.IsValid);
        var finding = Assert.Single(manifest.Findings);
        Assert.Equal("malformed_xml", finding.Code);
        Assert.Equal(VerificationFindingSeverity.Error, finding.Severity);
        Assert.Equal("/docProps/custom.xml", finding.Location?.EntryUri);
        var custom = Assert.Single(manifest.Entries, e => e.Uri == "/docProps/custom.xml");
        Assert.Equal("application/vnd.openxmlformats-officedocument.custom-properties+xml", custom.ContentType);
        Assert.True(custom.IsXml);
        Assert.NotNull(custom.RawBytesDigest);
        Assert.Null(custom.NormalizedXmlDigest);
        Assert.True(Assert.Single(manifest.Relationships, r => r.Id == "custom").IsTargetPresent);
        Assert.Equal(saved, bytes);
    }

    [Theory]
    [InlineData("text")]
    [InlineData("number")]
    [InlineData("boolean")]
    public void ValidCustomPropertiesRetainTheirXmlIdentity(string kind)
    {
        var manifest = PackageManifestGenerator.Generate(Package(kind));
        Assert.True(manifest.IsValid, string.Join("\n", manifest.Findings.Select(f => f.Message)));
        var custom = Assert.Single(manifest.Entries, e => e.Uri == "/docProps/custom.xml");
        Assert.True(custom.IsXml);
        Assert.NotNull(custom.RawBytesDigest);
        Assert.NotNull(custom.NormalizedXmlDigest);
        Assert.True(Assert.Single(manifest.Relationships, r => r.Id == "custom").IsTargetPresent);
    }

    [Fact]
    public void OmittedCustomPropertiesAreValid()
    {
        var manifest = PackageManifestGenerator.Generate(Package("none"));
        Assert.True(manifest.IsValid);
        Assert.DoesNotContain(manifest.Entries, e => e.Uri == "/docProps/custom.xml");
    }

    [Fact]
    public void CustomPropertiesStillObeyEntrySizeLimits()
    {
        var manifest = PackageManifestGenerator.Generate(Package("large"),
            new PackageManifestOptions { MaxEntryUncompressedBytes = 1024 });
        Assert.False(manifest.IsValid);
        Assert.Contains(manifest.Findings, f => f.Code == "entry_size_limit_exceeded"
            && f.Location?.EntryUri == "/docProps/custom.xml");
    }
}
