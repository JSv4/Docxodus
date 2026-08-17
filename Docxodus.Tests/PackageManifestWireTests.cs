// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.IO.Compression;
using System.Text;
using System.Text.Json;
using Docxodus;
using Docxodus.Verification;
using Xunit;

namespace OxPt;

public class PackageManifestWireTests
{
    [Fact]
    public void PM060_EntrySizes_AreLosslessDecimalStringsOnTheJsonWire()
    {
        const long beyondJavaScriptSafeInteger = 9_007_199_254_740_993L;
        var digest = new VerificationDigest
        {
            Algorithm = "SHA-256",
            Value = new string('0', 64),
        };
        var manifest = new PackageManifest
        {
            PackageKind = "zip",
            IsValid = true,
            RawPackageBytesDigest = digest,
            Entries =
            [
                new PackageManifestEntry
                {
                    Uri = "/large.bin",
                    Occurrence = 0,
                    ContentType = null,
                    ContentTypeSource = "unresolved",
                    Size = beyondJavaScriptSafeInteger,
                    CompressedSize = beyondJavaScriptSafeInteger - 1,
                    RawBytesDigest = digest,
                    NormalizedXmlDigest = null,
                    IsXml = false,
                    IsEncrypted = false,
                },
            ],
        };

        using var json = JsonDocument.Parse(manifest.ToJson());
        var entry = json.RootElement.GetProperty("entries")[0];
        Assert.Equal(JsonValueKind.String, entry.GetProperty("size").ValueKind);
        Assert.Equal("9007199254740993", entry.GetProperty("size").GetString());
        Assert.Equal("9007199254740992", entry.GetProperty("compressedSize").GetString());
    }

    [Fact]
    public void PM061_LiveManifestIncludesUnsavedXmlAndPreservesUndoRedo()
    {
        using var session = new DocxSession(DocxSession.CreateBlankDocxBytes());
        var paragraph = Assert.Single(session.FindByKind("p", "body")).Anchor.Id;
        var before = session.GetPackageManifest();
        var versionBeforeRead = session.Version;

        Assert.Equal(before.ToJson(), session.GetPackageManifest().ToJson());
        Assert.Equal(versionBeforeRead, session.Version);

        Assert.True(session.ReplaceText(paragraph, "unsaved manifest edit").Success);
        var after = session.GetPackageManifest();
        var versionAfterEdit = session.Version;
        Assert.NotEqual(before.NormalizedSemanticDigest, after.NormalizedSemanticDigest);
        Assert.Equal(versionAfterEdit, session.Version);

        Assert.True(session.Undo());
        Assert.Equal(before.NormalizedSemanticDigest,
            session.GetPackageManifest().NormalizedSemanticDigest);
        Assert.True(session.Redo());
        Assert.Equal(after.NormalizedSemanticDigest,
            session.GetPackageManifest().NormalizedSemanticDigest);
    }

    [Fact]
    public void PM062_MaximumXmlByteLimitDoesNotOverflowTheParserCharacterLimit()
    {
        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
        {
            WriteEntry(archive, "[Content_Types].xml", """
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="xml" ContentType="application/xml"/>
                </Types>
                """);
            WriteEntry(archive, "word/document.xml", "<document/>");
        }

        var manifest = PackageManifestGenerator.Generate(output.ToArray(),
            new PackageManifestOptions { MaxXmlPartBytes = long.MaxValue });

        Assert.True(manifest.IsValid,
            string.Join(Environment.NewLine, manifest.Findings.Select(item => item.Message)));
        Assert.DoesNotContain(manifest.Findings, item =>
            item.Code is "malformed_content_types" or "malformed_xml");
        Assert.All(manifest.Entries.Where(item => item.IsXml), item =>
            Assert.NotNull(item.NormalizedXmlDigest));
    }

    private static void WriteEntry(ZipArchive archive, string name, string value)
    {
        var entry = archive.CreateEntry(name, CompressionLevel.NoCompression);
        entry.LastWriteTime = new DateTimeOffset(2020, 1, 1, 0, 0, 0, TimeSpan.Zero);
        using var stream = entry.Open();
        stream.Write(Encoding.UTF8.GetBytes(value));
    }
}
