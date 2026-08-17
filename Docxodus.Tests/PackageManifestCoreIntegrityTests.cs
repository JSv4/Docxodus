// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Buffers.Binary;
using System.IO.Compression;
using System.Text;
using Docxodus.Verification;
using Xunit;

namespace OxPt;

public class PackageManifestCoreIntegrityTests
{
    [Fact]
    public void PM043_ReservedPercentEscapesRemainDistinctAndAreCanonicalized()
    {
        var package = BuildZip(
            ("[Content_Types].xml", Utf8(Types(
                "<Default Extension=\"bin\" ContentType=\"application/octet-stream\"/>"))),
            ("_rels/.rels", Utf8(Relationships(
                "<Relationship Id=\"rId1\" Type=\"urn:test\" Target=\"word/a%2c.bin\"/>"))),
            ("word/a%2c.bin", new byte[] { 1 }),
            ("word/a,.bin", new byte[] { 1 }));

        var manifest = PackageManifestGenerator.Generate(package);

        Assert.True(manifest.IsValid, string.Join("\n", manifest.Findings.Select(f => f.Message)));
        Assert.Contains(manifest.Entries, entry => entry.Uri == "/word/a%2C.bin");
        Assert.Contains(manifest.Entries, entry => entry.Uri == "/word/a,.bin");
        Assert.DoesNotContain(manifest.Findings, finding => finding.Code == "duplicate_entry");
        Assert.Contains(manifest.Relationships, relationship =>
            relationship.ResolvedTargetUri == "/word/a%2C.bin"
            && relationship.IsTargetPresent == true);
    }

    [Fact]
    public void PM044_OpcIriGrammarRejectsIllegalRawAndEncodedIunreservedCharacters()
    {
        var package = BuildZip(
            ("[Content_Types].xml", Utf8(Types(
                "<Default Extension=\"bin\" ContentType=\"application/octet-stream\"/>"))),
            ("word/raw space.bin", new byte[] { 1 }),
            ("word/[raw].bin", new byte[] { 2 }),
            ("word/%41.bin", new byte[] { 3 }),
            ("word/é.bin", new byte[] { 4 }));

        var manifest = PackageManifestGenerator.Generate(package);

        Assert.Equal(4, manifest.Findings.Count(finding => finding.Code == "unsafe_entry_path"));

        var logicalEscapes = PackageManifestGenerator.Generate(BuildZip(
            ("[Content_Types].xml", Utf8(Types(
                "<Default Extension=\"bin\" ContentType=\"application/octet-stream\"/>"
                + "<Override PartName=\"/word/%C3%A9.bin\" ContentType=\"application/octet-stream\"/>"))),
            ("_rels/.rels", Utf8(Relationships(
                "<Relationship Id=\"rId1\" Type=\"urn:test\" Target=\"word/%C3%A9.bin\"/>"))),
            ("word/%C3%A9.bin", new byte[] { 1 })));
        Assert.Contains(logicalEscapes.Findings, finding =>
            finding.Code == "invalid_content_type_part_name");
        Assert.Contains(logicalEscapes.Findings, finding =>
            finding.Code == "invalid_relationship_target");
    }

    [Fact]
    public void PM045_ZipUnicodeMappingIsLosslessAndOpcCaseFoldingIsAsciiOnly()
    {
        var package = BuildZip(
            ("[Content_Types].xml", Utf8(Types(
                "<Default Extension=\"bin\" ContentType=\"application/octet-stream\"/>"
                + "<Override PartName=\"/word/é.bin\" ContentType=\"application/octet-stream\"/>"))),
            ("_rels/.rels", Utf8(Relationships(
                "<Relationship Id=\"rId1\" Type=\"urn:test\" Target=\"word/é.bin\"/>"))),
            ("word/%c3%a9.bin", new byte[] { 1 }),
            ("word/%C3%89.bin", new byte[] { 2 }),
            ("word/%FC.bin", new byte[] { 3 }),
            ("word/%FF%C3%A9.bin", new byte[] { 4 }));

        var manifest = PackageManifestGenerator.Generate(package);

        Assert.True(manifest.IsValid, string.Join("\n", manifest.Findings.Select(f => f.Message)));
        Assert.DoesNotContain(manifest.Findings, finding =>
            finding.Code is "unsafe_entry_path" or "duplicate_entry" or "conflicting_entry");
        Assert.Contains(manifest.Entries, entry => entry.Uri == "/word/é.bin");
        Assert.Contains(manifest.Entries, entry => entry.Uri == "/word/É.bin");
        Assert.Contains(manifest.Entries, entry => entry.Uri == "/word/%FC.bin");
        Assert.Contains(manifest.Entries, entry => entry.Uri == "/word/%FFé.bin");
        Assert.Contains(manifest.Relationships, relationship =>
            relationship.ResolvedTargetUri == "/word/é.bin"
            && relationship.IsTargetPresent == true);

        var asciiCaseCollision = PackageManifestGenerator.Generate(BuildZip(
            ("[Content_Types].xml", Utf8(Types(
                "<Default Extension=\"bin\" ContentType=\"application/octet-stream\"/>"))),
            ("word/A.bin", new byte[] { 1 }),
            ("WORD/a.bin", new byte[] { 2 })));
        Assert.Contains(asciiCaseCollision.Findings, finding => finding.Code == "duplicate_entry");
        Assert.Contains(asciiCaseCollision.Findings, finding => finding.Code == "conflicting_entry");
    }

    [Fact]
    public void PM046_InterleavedPartNamesAreRejected()
    {
        var package = BuildZip(
            ("[Content_Types].xml", Utf8(Types(
                "<Default Extension=\"bin\" ContentType=\"application/octet-stream\"/>"
                + "<Override PartName=\"/word/a\" ContentType=\"application/octet-stream\"/>"))),
            ("word/a", new byte[] { 1 }),
            ("word/a/b.bin", new byte[] { 2 }));

        var manifest = PackageManifestGenerator.Generate(package);

        Assert.Contains(manifest.Findings, finding =>
            finding.Code == "interleaved_part_names"
            && finding.Location?.EntryUri == "/word/a/b.bin"
            && finding.Location.TargetUri == "/word/a");
    }

    [Fact]
    public void PM047_XmlDigestIgnoresDocumentWhitespaceAndExpandsQNameValues()
    {
        var first = XmlSemanticNormalizer.Parse(Utf8("""

            <root xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                  xmlns:t="urn:type" xsi:type="t:Thing"/>

            """), 10_000);
        var same = XmlSemanticNormalizer.Parse(Utf8("""<root xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:u="urn:type" xsi:type="u:Thing"/>"""), 10_000);
        var rebound = XmlSemanticNormalizer.Parse(Utf8("""<root xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:t="urn:other" xsi:type="t:Thing"/>"""), 10_000);

        var firstDigest = XmlSemanticNormalizer.Digest(first, "/custom.xml", false);
        var sameDigest = XmlSemanticNormalizer.Digest(same, "/custom.xml", false);
        var reboundDigest = XmlSemanticNormalizer.Digest(rebound, "/custom.xml", false);

        Assert.Equal(firstDigest, sameDigest);
        Assert.NotEqual(firstDigest, reboundDigest);
    }

    [Fact]
    public void PM048_AuthoritativeBinaryMimeOverridesXmlFileSuffix()
    {
        var package = BuildZip(
            ("[Content_Types].xml", Utf8(Types(
                "<Override PartName=\"/data/payload.xml\" ContentType=\"application/octet-stream\"/>"))),
            ("data/payload.xml", new byte[] { 0xff, 0x00, 0xfe }));

        var manifest = PackageManifestGenerator.Generate(package);
        var entry = Assert.Single(manifest.Entries,
            entry => entry.Uri == "/data/payload.xml");

        Assert.False(entry.IsXml);
        Assert.NotNull(entry.RawBytesDigest);
        Assert.Null(entry.NormalizedXmlDigest);
        Assert.DoesNotContain(manifest.Findings, finding => finding.Code == "malformed_xml");
    }

    [Fact]
    public void PM049_TxbxReferencesAndTargetModeCaseAreValidated()
    {
        var package = BuildZip(
            ("[Content_Types].xml", Utf8(Types(
                "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>"))),
            ("_rels/.rels", Utf8(Relationships(
                "<Relationship Id=\"rId1\" Type=\"urn:test\" Target=\"https://example.test\" TargetMode=\"external\"/>"))),
            ("word/document.xml", Utf8("""
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <w:body><w:p r:txbx="rIdTextbox"/></w:body>
                </w:document>
                """)));

        var manifest = PackageManifestGenerator.Generate(package);

        Assert.Contains(manifest.Findings, finding => finding.Code == "invalid_target_mode");
        var malformedRelationship = Assert.Single(manifest.Relationships);
        Assert.Equal("rId1", malformedRelationship.Id);
        Assert.Equal("Internal", malformedRelationship.TargetMode);
        Assert.Null(malformedRelationship.ResolvedTargetUri);
        Assert.Null(malformedRelationship.IsTargetPresent);
        Assert.Contains(manifest.Findings, finding =>
            finding.Code == "dangling_relationship"
            && finding.Location?.RelationshipId == "rIdTextbox");
    }

    [Fact]
    public void PM050_NonemptyTrailingSlashEntryParticipatesInIdentity()
    {
        var first = PackageManifestGenerator.Generate(DirectoryPayloadPackage(1));
        var second = PackageManifestGenerator.Generate(DirectoryPayloadPackage(2));

        Assert.Contains(first.Findings, finding => finding.Code == "nonempty_directory_entry");
        Assert.NotNull(first.OrderedOpcContentDigest);
        Assert.NotNull(first.NormalizedSemanticDigest);
        Assert.NotEqual(first.OrderedOpcContentDigest, second.OrderedOpcContentDigest);
        Assert.NotEqual(first.NormalizedSemanticDigest, second.NormalizedSemanticDigest);
    }

    [Fact]
    public void PM051_EntryCapUsesFullNameIndexForAbsenceAndPackageKind()
    {
        var package = BuildZip(
            ("[Content_Types].xml", Utf8(Types(
                "<Override PartName=\"/word/document.xml\" ContentType=\"application/xml\"/>"))),
            ("_rels/.rels", Utf8(Relationships(
                "<Relationship Id=\"rId1\" Type=\"urn:test\" Target=\"word/document.xml\"/>"))),
            ("word/document.xml", Utf8("<document/>")));

        var manifest = PackageManifestGenerator.Generate(package,
            new PackageManifestOptions { MaxEntryCount = 2 });

        Assert.Equal("opc", manifest.PackageKind);
        Assert.Contains(manifest.Findings, finding => finding.Code == "entry_count_limit_exceeded");
        Assert.DoesNotContain(manifest.Findings, finding =>
            finding.Code is "missing_content_types" or "missing_content_type_target"
                or "missing_target" or "missing_relationship_owner");
        Assert.Contains(manifest.Relationships, relationship =>
            relationship.ResolvedTargetUri == "/word/document.xml"
            && relationship.IsTargetPresent == true);
    }

    [Fact]
    public void PM052_StrictOoxmlDoesNotInventStrictOpcMetadataNamespaces()
    {
        var package = BuildZip(
            ("[Content_Types].xml", Utf8("""
                <Types xmlns="http://purl.oclc.org/ooxml/package/content-types">
                  <Default Extension="xml" ContentType="application/xml"/>
                </Types>
                """)),
            ("_rels/.rels", Utf8("""
                <Relationships xmlns="http://purl.oclc.org/ooxml/package/relationships">
                  <Relationship Id="rId1" Type="urn:test" Target="word/document.xml"/>
                </Relationships>
                """)),
            ("word/document.xml", Utf8("<document/>")));

        var manifest = PackageManifestGenerator.Generate(package);

        Assert.Contains(manifest.Findings, finding => finding.Code == "malformed_content_types");
        Assert.Contains(manifest.Findings, finding => finding.Code == "malformed_relationship_part");
    }

    [Fact]
    public void PM053_MediaFactsCountImageAudioAndVideoMimeFamilies()
    {
        var package = BuildZip(
            ("[Content_Types].xml", Utf8(Types(
                "<Default Extension=\"png\" ContentType=\"image/png; profile=screen\"/>"
                + "<Default Extension=\"mp3\" ContentType=\"audio/mpeg; codecs=mp3\"/>"
                + "<Default Extension=\"mp4\" ContentType=\"video/mp4; codecs=avc1\"/>"))),
            ("word/media/image.png", new byte[] { 1 }),
            ("word/media/audio.mp3", new byte[] { 2 }),
            ("word/media/video.mp4", new byte[] { 3 }));

        var manifest = PackageManifestGenerator.Generate(package);

        Assert.Equal(3, manifest.Facts.MediaPartCount);
    }

    [Fact]
    public void PM054_CentralDirectoryCrcMismatchIsReported()
    {
        var package = BuildZip(
            ("[Content_Types].xml", Utf8(Types(
                "<Default Extension=\"bin\" ContentType=\"application/octet-stream\"/>"))),
            ("data/payload.bin", new byte[] { 1, 2, 3 }));
        var central = FindSignature(package, 0x02014b50);
        Assert.True(central >= 0);
        var expected = BinaryPrimitives.ReadUInt32LittleEndian(package.AsSpan(central + 16, 4));
        BinaryPrimitives.WriteUInt32LittleEndian(package.AsSpan(central + 16, 4), expected ^ 1);

        var manifest = PackageManifestGenerator.Generate(package);

        Assert.Contains(manifest.Findings, finding => finding.Code == "crc_mismatch");
    }

    [Fact]
    public void PM055_ParameterizedKnownXmlUsesItsMediaTypeEssence()
    {
        const string contentType =
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml ; charset=utf-8";
        var contentTypes = Utf8(Types(
            $"<Override PartName=\"/word/document.dat\" ContentType=\"{contentType}\"/>"));
        var compact = BuildZip(
            ("[Content_Types].xml", contentTypes),
            ("word/document.dat", Utf8("""
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>
                """)));
        var pretty = BuildZip(
            ("[Content_Types].xml", contentTypes),
            ("word/document.dat", Utf8("""
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                  <w:body>
                    <w:p></w:p>
                  </w:body>
                </w:document>
                """)));

        var compactManifest = PackageManifestGenerator.Generate(compact);
        var prettyManifest = PackageManifestGenerator.Generate(pretty);
        var compactEntry = compactManifest.Entries.Single(entry =>
            entry.Uri == "/word/document.dat");
        var prettyEntry = prettyManifest.Entries.Single(entry =>
            entry.Uri == "/word/document.dat");

        Assert.True(compactEntry.IsXml);
        Assert.NotNull(compactEntry.NormalizedXmlDigest);
        Assert.Equal(compactEntry.NormalizedXmlDigest, prettyEntry.NormalizedXmlDigest);
        Assert.NotEqual(compactManifest.OrderedOpcContentDigest,
            prettyManifest.OrderedOpcContentDigest);
        Assert.Equal(compactManifest.NormalizedSemanticDigest,
            prettyManifest.NormalizedSemanticDigest);
        Assert.Equal(1, compactManifest.Facts.ParagraphCount);
        Assert.Equal(1, prettyManifest.Facts.ParagraphCount);
    }

    [Fact]
    public void PM056_OpcMetadataSortUsesAsciiKeysAndCompleteAttributeIdentity()
    {
        var unicodeFirst = XmlSemanticNormalizer.Parse(Utf8(Types(
            "<Default Extension=\"É\" ContentType=\"application/xml\"/>"
            + "<Default Extension=\"é\" ContentType=\"application/xml\"/>")), 10_000);
        var unicodeReordered = XmlSemanticNormalizer.Parse(Utf8(Types(
            "<Default Extension=\"é\" ContentType=\"application/xml\"/>"
            + "<Default Extension=\"É\" ContentType=\"application/xml\"/>")), 10_000);
        var asciiTieFirst = XmlSemanticNormalizer.Parse(Utf8(Types(
            "<Default Extension=\"XML\" ContentType=\"application/xml\"/>"
            + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>")), 10_000);
        var asciiTieReordered = XmlSemanticNormalizer.Parse(Utf8(Types(
            "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
            + "<Default Extension=\"XML\" ContentType=\"application/xml\"/>")), 10_000);
        var namespacedAttributeFirst = XmlSemanticNormalizer.Parse(Utf8(Relationships("""
            <Relationship xmlns:x="urn:extension" x:Id="z" Id="a" Type="urn:a" Target="a"/>
            <Relationship xmlns:x="urn:extension" x:Id="a" Id="z" Type="urn:z" Target="z"/>
            """)), 10_000);
        var namespacedAttributeReordered = XmlSemanticNormalizer.Parse(Utf8(Relationships("""
            <Relationship xmlns:x="urn:extension" Id="z" x:Id="a" Type="urn:z" Target="z"/>
            <Relationship xmlns:x="urn:extension" Id="a" x:Id="z" Type="urn:a" Target="a"/>
            """)), 10_000);
        var completeAttributeFirst = XmlSemanticNormalizer.Parse(Utf8(Relationships("""
            <Relationship xmlns:x="urn:extension" Id="same" Type="urn:same" Target="same" x:Flag="b"/>
            <Relationship xmlns:x="urn:extension" Id="same" Type="urn:same" Target="same" x:Flag="a"/>
            """)), 10_000);
        var completeAttributeReordered = XmlSemanticNormalizer.Parse(Utf8(Relationships("""
            <Relationship xmlns:x="urn:extension" x:Flag="a" Target="same" Type="urn:same" Id="same"/>
            <Relationship xmlns:x="urn:extension" x:Flag="b" Target="same" Type="urn:same" Id="same"/>
            """)), 10_000);

        Assert.Equal(
            XmlSemanticNormalizer.Digest(unicodeFirst, "/[Content_Types].xml", true),
            XmlSemanticNormalizer.Digest(unicodeReordered, "/[Content_Types].xml", true));
        Assert.Equal(
            XmlSemanticNormalizer.Digest(asciiTieFirst, "/[Content_Types].xml", true),
            XmlSemanticNormalizer.Digest(asciiTieReordered, "/[Content_Types].xml", true));
        Assert.Equal(
            XmlSemanticNormalizer.Digest(namespacedAttributeFirst, "/_rels/.rels", true),
            XmlSemanticNormalizer.Digest(namespacedAttributeReordered, "/_rels/.rels", true));
        Assert.Equal(
            XmlSemanticNormalizer.Digest(completeAttributeFirst, "/_rels/.rels", true),
            XmlSemanticNormalizer.Digest(completeAttributeReordered, "/_rels/.rels", true));
    }

    [Fact]
    public void PM057_CaseEquivalentDuplicateOccurrencesAreStableAcrossRepacking()
    {
        var metadata = ("[Content_Types].xml", Utf8(Types(
            "<Default Extension=\"bin\" ContentType=\"application/octet-stream\"/>")));
        var upper = ("word/A.bin", new byte[] { 1 });
        var lower = ("WORD/a.bin", new byte[] { 1 });

        var first = PackageManifestGenerator.Generate(BuildZip(metadata, upper, lower));
        var reordered = PackageManifestGenerator.Generate(BuildZip(metadata, lower, upper));

        Assert.Contains(first.Findings, finding => finding.Code == "duplicate_entry");
        Assert.Equal(first.OrderedOpcContentDigest, reordered.OrderedOpcContentDigest);
        Assert.Equal(first.NormalizedSemanticDigest, reordered.NormalizedSemanticDigest);
        Assert.Equal(
            first.Entries.Select(entry => (entry.Uri, entry.Occurrence)),
            reordered.Entries.Select(entry => (entry.Uri, entry.Occurrence)));
    }

    private static byte[] DirectoryPayloadPackage(byte payload) => BuildZip(
        ("[Content_Types].xml", Utf8(Types(string.Empty))),
        ("word/", new[] { payload }));

    private static string Types(string declarations) => $$"""
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          {{declarations}}
        </Types>
        """;

    private static string Relationships(string declarations) => $$"""
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          {{declarations}}
        </Relationships>
        """;

    private static byte[] BuildZip(params (string Name, byte[] Bytes)[] entries)
    {
        using var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var (name, bytes) in entries)
            {
                var entry = archive.CreateEntry(name, CompressionLevel.NoCompression);
                entry.LastWriteTime = new DateTimeOffset(2024, 1, 1, 0, 0, 0, TimeSpan.Zero);
                using var output = entry.Open();
                output.Write(bytes);
            }
        }
        return stream.ToArray();
    }

    private static byte[] Utf8(string value) => Encoding.UTF8.GetBytes(value);

    private static int FindSignature(byte[] bytes, uint signature)
    {
        for (var index = 0; index <= bytes.Length - sizeof(uint); index++)
            if (BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(index, sizeof(uint))) == signature)
                return index;
        return -1;
    }
}
