// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Buffers.Binary;
using System.IO.Compression;
using System.Text;
using System.Text.Json;
using Docxodus.Internal;
using Docxodus.McpServer;
using Docxodus.Verification;
using Xunit;

namespace OxPt;

public class PackageManifestTests
{
    [Fact]
    public void PM001_RichPackage_CapturesEntriesRelationshipsAndRendererFacts()
    {
        var manifest = PackageManifestGenerator.Generate(BuildRichPackage());

        Assert.True(manifest.IsValid, string.Join("\n", manifest.Findings.Select(f => f.Message)));
        Assert.Equal("opc", manifest.PackageKind);
        Assert.NotNull(manifest.OrderedOpcContentDigest);
        Assert.NotNull(manifest.NormalizedSemanticDigest);
        Assert.All(manifest.Entries, entry =>
        {
            Assert.NotNull(entry.ContentType);
            Assert.Equal(64, entry.RawBytesDigest?.Value.Length);
            if (entry.IsXml)
                Assert.Equal(64, entry.NormalizedXmlDigest?.Value.Length);
        });
        Assert.Contains(manifest.Relationships, relationship =>
            relationship.OwnerUri == "/"
            && relationship.Id == "rIdMain"
            && relationship.ResolvedTargetUri == "/word/document.xml"
            && relationship.IsTargetPresent == true);
        Assert.Contains(manifest.Relationships, relationship =>
            relationship.Id == "rIdExternal"
            && relationship.TargetMode == "External"
            && relationship.ResolvedTargetUri is null
            && relationship.IsTargetPresent is null);
        Assert.Equal("/word/document.xml", manifest.Facts.MainDocumentUri);
        Assert.False(manifest.Facts.IsStrictOoxml);
        Assert.Equal(1, manifest.Facts.SectionCount);
        Assert.True(manifest.Facts.ParagraphCount >= 8);
        Assert.Equal(1, manifest.Facts.TableCount);
        Assert.Equal(1, manifest.Facts.HeaderPartCount);
        Assert.Equal(1, manifest.Facts.FooterPartCount);
        Assert.Equal(1, manifest.Facts.FootnoteCount);
        Assert.Equal(1, manifest.Facts.EndnoteCount);
        Assert.Equal(2, manifest.Facts.StyleCount);
        Assert.Equal(2, manifest.Facts.NumberingDefinitionCount);
        Assert.Equal(1, manifest.Facts.ThemePartCount);
        Assert.Equal(1, manifest.Facts.MediaPartCount);
        Assert.Equal(2, manifest.Facts.CustomXmlPartCount);
        Assert.Equal(1, manifest.Facts.DrawingCount);
        Assert.Equal(1, manifest.Facts.AltChunkCount);
        Assert.Equal(1, manifest.Facts.FieldCount);
        Assert.Equal(5, manifest.Facts.Revisions.Total);
        Assert.Equal(2, manifest.Facts.Annotations.Comments);
        Assert.Equal(1, manifest.Facts.Annotations.CommentReplies);
        Assert.Equal(2, manifest.Facts.Annotations.ThreadedCommentMetadata);
        Assert.Equal(1, manifest.Facts.Annotations.ResolvedComments);
        Assert.Equal(1, manifest.Facts.Annotations.People);
        Assert.Equal(1, manifest.Facts.Annotations.DocxodusAnnotations);
        Assert.Contains(manifest.Entries, entry =>
            entry.Uri == "/word/data/payload.weird"
            && entry.ContentType == "application/x-docxodus-test");
    }

    [Fact]
    public void PM002_Generation_IsDeterministicAndDoesNotMutateInput()
    {
        var bytes = BuildRichPackage();
        var before = bytes.ToArray();

        var first = PackageManifestGenerator.Generate(bytes);
        var second = PackageManifestGenerator.Generate(bytes);

        Assert.Equal(before, bytes);
        Assert.Equal(first.RawPackageBytesDigest, second.RawPackageBytesDigest);
        Assert.Equal(first.OrderedOpcContentDigest, second.OrderedOpcContentDigest);
        Assert.Equal(first.NormalizedSemanticDigest, second.NormalizedSemanticDigest);
        Assert.Equal(first.ToJson(), second.ToJson());
        using var parsed = JsonDocument.Parse(first.ToJson());
        Assert.Equal(PackageManifest.SchemaId, parsed.RootElement.GetProperty("schema").GetString());
        Assert.Equal("SHA-256", parsed.RootElement
            .GetProperty("rawPackageBytesDigest").GetProperty("algorithm").GetString());
    }

    [Fact]
    public void PM003_RepackChangesOnlyRawPackageBytesDigest()
    {
        var first = BuildRichPackage(reverseEntries: false,
            compression: CompressionLevel.Optimal,
            timestamp: new DateTimeOffset(2024, 1, 1, 0, 0, 0, TimeSpan.Zero));
        var repacked = BuildRichPackage(reverseEntries: true,
            compression: CompressionLevel.NoCompression,
            timestamp: new DateTimeOffset(2025, 1, 1, 0, 0, 0, TimeSpan.Zero));

        var left = PackageManifestGenerator.Generate(first);
        var right = PackageManifestGenerator.Generate(repacked);

        Assert.NotEqual(left.RawPackageBytesDigest, right.RawPackageBytesDigest);
        Assert.Equal(left.OrderedOpcContentDigest, right.OrderedOpcContentDigest);
        Assert.Equal(left.NormalizedSemanticDigest, right.NormalizedSemanticDigest);
    }

    [Fact]
    public void PM004_KnownOoxmlPrettyPrinting_IsSerializationOnly()
    {
        var compact = BuildRichPackage();
        var pretty = RewriteEntry(compact, "word/document.xml", xml =>
            xml.Replace("><", ">\n  <", StringComparison.Ordinal));

        var left = PackageManifestGenerator.Generate(compact);
        var right = PackageManifestGenerator.Generate(pretty);

        Assert.NotEqual(left.OrderedOpcContentDigest, right.OrderedOpcContentDigest);
        Assert.Equal(left.Entries.Single(e => e.Uri == "/word/document.xml").NormalizedXmlDigest,
            right.Entries.Single(e => e.Uri == "/word/document.xml").NormalizedXmlDigest);
        Assert.Equal(left.NormalizedSemanticDigest, right.NormalizedSemanticDigest);
    }

    [Fact]
    public void PM005_OpaqueCustomXmlWhitespace_RemainsSemantic()
    {
        var compact = BuildRichPackage(customXml: "<data xmlns=\"urn:opaque\"><a>one</a><b>two</b></data>");
        var spaced = BuildRichPackage(customXml: "<data xmlns=\"urn:opaque\"><a>one</a> <b>two</b></data>");

        var left = PackageManifestGenerator.Generate(compact);
        var right = PackageManifestGenerator.Generate(spaced);

        Assert.NotEqual(left.Entries.Single(e => e.Uri == "/customXml/item2.xml").NormalizedXmlDigest,
            right.Entries.Single(e => e.Uri == "/customXml/item2.xml").NormalizedXmlDigest);
        Assert.NotEqual(left.NormalizedSemanticDigest, right.NormalizedSemanticDigest);
    }

    [Fact]
    public void PM006_XmlSpacePreserve_KeepsWhitespaceInKnownOoxml()
    {
        var first = BuildRichPackage(documentText: "<w:t xml:space=\"preserve\"> </w:t>");
        var second = BuildRichPackage(documentText: "<w:t xml:space=\"preserve\">  </w:t>");

        Assert.NotEqual(PackageManifestGenerator.Generate(first).NormalizedSemanticDigest,
            PackageManifestGenerator.Generate(second).NormalizedSemanticDigest);
    }

    [Fact]
    public void PM007_SemanticXmlChange_ChangesContentAndSemanticDigests()
    {
        var first = PackageManifestGenerator.Generate(BuildRichPackage(documentText: "<w:t>alpha</w:t>"));
        var second = PackageManifestGenerator.Generate(BuildRichPackage(documentText: "<w:t>beta</w:t>"));

        Assert.NotEqual(first.RawPackageBytesDigest, second.RawPackageBytesDigest);
        Assert.NotEqual(first.OrderedOpcContentDigest, second.OrderedOpcContentDigest);
        Assert.NotEqual(first.NormalizedSemanticDigest, second.NormalizedSemanticDigest);
    }

    [Fact]
    public void PM008_StrictOoxml_ParsesNormalizesAndCountsWithoutConflatingConformanceClasses()
    {
        var strictBytes = BuildRichPackage(strict: true);
        var strictFirst = PackageManifestGenerator.Generate(strictBytes);
        var strictSecond = PackageManifestGenerator.Generate(strictBytes);
        var transitional = PackageManifestGenerator.Generate(BuildRichPackage());

        Assert.True(strictFirst.IsValid, string.Join("\n", strictFirst.Findings.Select(f => f.Message)));
        Assert.True(strictFirst.Facts.IsStrictOoxml);
        Assert.Equal(1, strictFirst.Facts.SectionCount);
        Assert.Equal(5, strictFirst.Facts.Revisions.Total);
        Assert.Equal(strictFirst.ToJson(), strictSecond.ToJson());
        Assert.NotEqual(strictFirst.NormalizedSemanticDigest, transitional.NormalizedSemanticDigest);
    }

    [Fact]
    public void PM009_DuplicateAndConflictingEntries_ArePreservedAndReported()
    {
        var entries = MinimalEntries().ToList();
        entries.Add(("word/duplicate.bin", new byte[] { 1 }));
        entries.Add(("word/duplicate.bin", new byte[] { 2 }));

        var manifest = PackageManifestGenerator.Generate(
            BuildZip(entries, CompressionLevel.NoCompression, DefaultTimestamp));

        Assert.False(manifest.IsValid);
        var duplicates = manifest.Entries.Where(e => e.Uri == "/word/duplicate.bin").ToList();
        Assert.Equal(2, duplicates.Count);
        Assert.Equal(new[] { 0, 1 }, duplicates.Select(e => e.Occurrence));
        Assert.Contains(manifest.Findings, finding => finding.Code == "duplicate_entry");
        Assert.Contains(manifest.Findings, finding => finding.Code == "conflicting_entry");
    }

    [Fact]
    public void PM010_ContentTypeDuplicatesConflictsAndMissingOverrideTargets_AreStructured()
    {
        const string contentTypes = """
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="xml" ContentType="application/xml"/>
              <Default Extension="xml" ContentType="application/xml"/>
              <Default Extension="xml" ContentType="text/xml"/>
              <Override PartName="/missing.xml" ContentType="application/xml"/>
            </Types>
            """;

        var manifest = PackageManifestGenerator.Generate(BuildZip(
            MinimalEntries(contentTypes), CompressionLevel.NoCompression, DefaultTimestamp));

        Assert.Contains(manifest.Findings, finding => finding.Code == "duplicate_content_type");
        Assert.Contains(manifest.Findings, finding => finding.Code == "conflicting_content_type");
        Assert.Contains(manifest.Findings, finding => finding.Code == "missing_content_type_target");
    }

    [Fact]
    public void PM011_MissingTargetsAndDanglingRelationshipIds_AreDistinctFindings()
    {
        const string document = """
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <w:body><w:p><w:hyperlink r:id="rIdAbsent"><w:r><w:t>x</w:t></w:r></w:hyperlink></w:p></w:body>
            </w:document>
            """;
        const string rels = """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rIdMissing" Type="urn:test" Target="media/missing.png"/>
            </Relationships>
            """;
        var entries = MinimalEntries().ToList();
        Replace(entries, "word/document.xml", Utf8(document));
        entries.Add(("word/_rels/document.xml.rels", Utf8(rels)));

        var manifest = PackageManifestGenerator.Generate(
            BuildZip(entries, CompressionLevel.NoCompression, DefaultTimestamp));

        Assert.Contains(manifest.Findings, finding => finding.Code == "missing_target"
            && finding.Location?.RelationshipId == "rIdMissing");
        Assert.Contains(manifest.Findings, finding => finding.Code == "dangling_relationship"
            && finding.Location?.RelationshipId == "rIdAbsent");
    }

    [Fact]
    public void PM012_MalformedOleEncryptionAndZipEncryption_ReturnStructuredManifests()
    {
        var malformed = PackageManifestGenerator.Generate(Utf8("not a zip"));
        var ole = new byte[512];
        new byte[] { 0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1 }.CopyTo(ole, 0);
        Encoding.Unicode.GetBytes("EncryptedPackage").CopyTo(ole, 64);
        var corruptOle = PackageManifestGenerator.Generate(ole);
        var encryptedZip = PackageManifestGenerator.Generate(MarkFirstEntryEncrypted(BuildZip(
            MinimalEntries(), CompressionLevel.NoCompression, DefaultTimestamp)));

        Assert.Equal("malformed", malformed.PackageKind);
        Assert.Contains(malformed.Findings, finding => finding.Code == "malformed_package");
        Assert.Equal("ole", corruptOle.PackageKind);
        Assert.Contains(corruptOle.Findings, finding => finding.Code == "unsupported_compound_file");
        Assert.Equal("zip-encrypted", encryptedZip.PackageKind);
        Assert.Contains(encryptedZip.Findings, finding => finding.Code == "unsupported_zip_encryption");
        Assert.Contains(encryptedZip.Entries, entry => entry.IsEncrypted == true);
    }

    [Fact]
    public void PM013_UnsafePathsDtdAndSafetyLimits_AreBoundedFindings()
    {
        var entries = MinimalEntries().ToList();
        entries.Add(("../escape.xml", Utf8("<!DOCTYPE x [<!ENTITY y 'boom'>]><x>&y;</x>")));
        var bytes = BuildZip(entries, CompressionLevel.Optimal, DefaultTimestamp);
        var manifest = PackageManifestGenerator.Generate(bytes, new PackageManifestOptions
        {
            MaxEntryCount = 100,
            MaxTotalUncompressedBytes = 1024 * 1024,
            MaxXmlPartBytes = 1024 * 1024,
            MaxCompressionRatio = 10_000,
            MaxUriLength = 2_048,
        });

        Assert.Contains(manifest.Findings, finding => finding.Code == "unsafe_entry_path");
        Assert.Contains(manifest.Findings, finding => finding.Code == "malformed_xml");

        var totalLimited = PackageManifestGenerator.Generate(bytes, new PackageManifestOptions
        {
            MaxEntryCount = 100,
            MaxTotalUncompressedBytes = 1,
            MaxXmlPartBytes = 1024,
            MaxCompressionRatio = 10_000,
            MaxUriLength = 2_048,
        });
        Assert.Contains(totalLimited.Findings,
            finding => finding.Code == "total_expansion_limit_exceeded");
        Assert.Null(totalLimited.OrderedOpcContentDigest);
    }

    [Fact]
    public void PM014_ActualExpansionBudget_IsSharedAcrossEntries()
    {
        var entries = MinimalEntries().ToList();
        entries.Add(("word/data/a.bin", Enumerable.Repeat((byte)0xa1, 8).ToArray()));
        entries.Add(("word/data/b.bin", Enumerable.Repeat((byte)0xb2, 8).ToArray()));
        var actualTotal = entries.Sum(entry => (long)entry.Bytes.Length);
        var package = RewriteCentralUncompressedSizes(
            BuildZip(entries, CompressionLevel.NoCompression, DefaultTimestamp),
            new Dictionary<string, uint>
            {
                ["word/data/a.bin"] = 1,
                ["word/data/b.bin"] = 1,
            });

        var manifest = PackageManifestGenerator.Generate(package, new PackageManifestOptions
        {
            MaxEntryCount = 100,
            MaxTotalUncompressedBytes = actualTotal - 7,
            MaxXmlPartBytes = 1024 * 1024,
            MaxCompressionRatio = 10_000,
            MaxUriLength = 2_048,
        });

        Assert.DoesNotContain(manifest.Findings,
            finding => finding.Code == "total_expansion_limit_exceeded");
        Assert.Contains(manifest.Findings,
            finding => finding.Code == "entry_expansion_limit_exceeded");
        Assert.Null(manifest.OrderedOpcContentDigest);
    }

    [Fact]
    public void PM015_LiveSessionManifest_IsRepeatableAndReadOnly()
    {
        var handle = DocxSessionOps.OpenSession(DocxSessionOps.CreateBlankDocx(), settings: null);
        try
        {
            var version = DocxSessionOps.GetVersion(handle);
            var packageHash = DocxSessionOps.GetPackageContentHash(handle);
            var first = VerificationOps.GetPackageManifest(handle);
            var second = VerificationOps.GetPackageManifest(handle);

            Assert.Equal(first, second);
            Assert.Equal(version, DocxSessionOps.GetVersion(handle));
            Assert.Equal(packageHash, DocxSessionOps.GetPackageContentHash(handle));
            using var parsed = JsonDocument.Parse(first);
            Assert.Equal(PackageManifest.SchemaId,
                parsed.RootElement.GetProperty("schema").GetString());
            Assert.Equal("opc", parsed.RootElement.GetProperty("packageKind").GetString());
        }
        finally
        {
            DocxSessionOps.CloseSession(handle);
        }
    }

    [Fact]
    public void PM016_McpCatalog_AdvertisesFullPackageManifestWithoutAnchorScope()
    {
        var tool = Assert.Single(ToolCatalog.Tools,
            definition => definition.Name == "docxodus_get_content");
        using var schema = JsonDocument.Parse(tool.InputSchemaJson);
        var properties = schema.RootElement.GetProperty("properties");
        Assert.Contains(properties.GetProperty("format").GetProperty("enum").EnumerateArray(),
            value => value.GetString() == "manifest");
        Assert.Contains("complete package",
            properties.GetProperty("anchorId").GetProperty("description").GetString());
    }

    [Fact]
    public void PM017_RawOpcSegmentGrammar_IsValidatedBeforePercentDecoding()
    {
        var entries = MinimalEntries().ToList();
        entries.Add(("word/encoded%2fslash.xml", Utf8("<x/>")));
        entries.Add(("word/encoded%5Cbackslash.xml", Utf8("<x/>")));
        entries.Add(("word/encoded%41unreserved.xml", Utf8("<x/>")));
        entries.Add(("word/%2e%2e/encoded-dot.xml", Utf8("<x/>")));
        entries.Add(("word/malformed%GG.xml", Utf8("<x/>")));
        entries.Add(("word//empty.xml", Utf8("<x/>")));
        entries.Add(("word/trailing./entry.xml", Utf8("<x/>")));

        var manifest = PackageManifestGenerator.Generate(
            BuildZip(entries, CompressionLevel.NoCompression, DefaultTimestamp));

        Assert.True(manifest.Findings.Count(finding => finding.Code == "unsafe_entry_path") >= 7);

        var invalidOverride = RewriteEntry(
            BuildZip(MinimalEntries(), CompressionLevel.NoCompression, DefaultTimestamp),
            "[Content_Types].xml",
            xml => xml.Replace("/word/document.xml", "/%77ord/document.xml",
                StringComparison.Ordinal));
        Assert.Contains(PackageManifestGenerator.Generate(invalidOverride).Findings,
            finding => finding.Code == "invalid_content_type_part_name");

        var invalidTarget = RewriteEntry(
            BuildZip(MinimalEntries(), CompressionLevel.NoCompression, DefaultTimestamp),
            "_rels/.rels",
            xml => xml.Replace("word/document.xml", "word/%2edocument.xml",
                StringComparison.Ordinal));
        Assert.Contains(PackageManifestGenerator.Generate(invalidTarget).Findings,
            finding => finding.Code == "invalid_relationship_target");
    }

    [Fact]
    public void PM018_ActualPerEntryExpansionCeiling_AppliesToContentTypesPreload()
    {
        var package = RewriteCentralUncompressedSizes(
            BuildZip(MinimalEntries(), CompressionLevel.NoCompression, DefaultTimestamp),
            new Dictionary<string, uint> { ["[Content_Types].xml"] = 1 });

        var manifest = PackageManifestGenerator.Generate(package, new PackageManifestOptions
        {
            MaxEntryCount = 100,
            MaxTotalUncompressedBytes = 1024 * 1024,
            MaxXmlPartBytes = 1024 * 1024,
            MaxCompressionRatio = 0.5,
            MaxUriLength = 2_048,
        });

        Assert.Contains(manifest.Findings, finding =>
            finding.Code == "compression_ratio_limit_exceeded"
            && finding.Location?.EntryUri == "/[Content_Types].xml");
        Assert.Null(manifest.OrderedOpcContentDigest);
    }

    [Fact]
    public void PM019_RevisionTotals_CoverStructuralAndCustomXmlRangeFamilies()
    {
        var package = BuildRichPackage(documentText: """
            <w:cellIns/><w:cellDel/><w:cellMerge/>
            <w:customXmlInsRangeStart/><w:customXmlInsRangeEnd/>
            <w:customXmlDelRangeStart/><w:customXmlDelRangeEnd/>
            <w:customXmlMoveFromRangeStart/><w:customXmlMoveFromRangeEnd/>
            <w:customXmlMoveToRangeStart/><w:customXmlMoveToRangeEnd/>
            """);

        var manifest = PackageManifestGenerator.Generate(package);
        var revisions = manifest.Facts.Revisions;

        Assert.Equal(3, revisions.StructuralChanges);
        Assert.Equal(4, revisions.OtherChanges);
        Assert.Equal(revisions.Insertions + revisions.Deletions + revisions.MoveFrom
            + revisions.MoveTo + revisions.PropertyChanges + revisions.StructuralChanges
            + revisions.OtherChanges, revisions.Total);
        using var json = JsonDocument.Parse(manifest.ToJson());
        var revisionJson = json.RootElement.GetProperty("facts").GetProperty("revisions");
        Assert.Equal(3, revisionJson.GetProperty("structuralChanges").GetInt32());
        Assert.Equal(4, revisionJson.GetProperty("otherChanges").GetInt32());
    }

    [Fact]
    public void PM020_DanglingReferencesAndFacts_RequireExactNamespacesAndPartTypes()
    {
        var baseline = PackageManifestGenerator.Generate(BuildRichPackage());
        var contaminatedBytes = BuildRichPackage(customXml: """
            <data xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                  xmlns:fake="urn:fake:relationships"
                  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <w:p><w:ins/><w:tbl/></w:p>
              <item fake:id="rIdFake" r:unsupported="rIdAlsoFake"/>
            </data>
            """);
        contaminatedBytes = RewriteEntry(contaminatedBytes, "word/styles.xml", xml =>
            xml.Replace("</w:styles>", "<w:p><w:ins/><w:tbl/></w:p></w:styles>",
                StringComparison.Ordinal));
        var contaminated = PackageManifestGenerator.Generate(contaminatedBytes);

        Assert.Equal(baseline.Facts.ParagraphCount, contaminated.Facts.ParagraphCount);
        Assert.Equal(baseline.Facts.TableCount, contaminated.Facts.TableCount);
        Assert.Equal(baseline.Facts.Revisions, contaminated.Facts.Revisions);
        Assert.DoesNotContain(contaminated.Findings, finding =>
            finding.Code == "dangling_relationship"
            && finding.Location?.RelationshipId is "rIdFake" or "rIdAlsoFake");
    }

    [Fact]
    public void PM021_UnavailableEncryptionDetection_IsExplicitAndNeverFalse()
    {
        var valid = BuildZip(MinimalEntries(), CompressionLevel.NoCompression, DefaultTimestamp);
        var withTrailingGarbage = valid.Concat(new byte[] { 0xa5 }).ToArray();

        var manifest = PackageManifestGenerator.Generate(withTrailingGarbage);

        Assert.Contains(manifest.Findings,
            finding => finding.Code == "zip_encryption_detection_unavailable");
        Assert.All(manifest.Entries, entry => Assert.Null(entry.IsEncrypted));
        Assert.All(manifest.Entries, entry => Assert.Null(entry.RawBytesDigest));
        using var json = JsonDocument.Parse(manifest.ToJson());
        Assert.All(json.RootElement.GetProperty("entries").EnumerateArray(), entry =>
            Assert.Equal(JsonValueKind.Null, entry.GetProperty("isEncrypted").ValueKind));
        Assert.Null(manifest.OrderedOpcContentDigest);

        // Not reading [Content_Types].xml is one failure, not one failure per part.
        Assert.DoesNotContain(manifest.Findings, finding => finding.Code == "missing_content_type");
        Assert.Single(manifest.Findings, finding => finding.Code == "content_types_unreadable");
    }

    [Fact]
    public void PM022_Zip64CentralDirectoryEncryptionFlags_AreParsedAuthoritatively()
    {
        var zip64 = PromoteToZip64Directory(BuildZip(
            MinimalEntries(), CompressionLevel.NoCompression, DefaultTimestamp));

        var manifest = PackageManifestGenerator.Generate(zip64);

        Assert.DoesNotContain(manifest.Findings,
            finding => finding.Code == "zip_encryption_detection_unavailable");
        Assert.All(manifest.Entries, entry => Assert.False(entry.IsEncrypted));
        Assert.NotNull(manifest.OrderedOpcContentDigest);
    }

    [Fact]
    public void PM023_UnknownVendorExtensionXml_PreservesFormattingWhitespace()
    {
        const string contentTypes = """
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="xml" ContentType="application/xml"/>
              <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
              <Override PartName="/word/opaque.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.unknownExtension+xml"/>
            </Types>
            """;
        var compactEntries = MinimalEntries(contentTypes).ToList();
        compactEntries.Add(("word/opaque.xml", Utf8("<opaque><a/><b/></opaque>")));
        var spacedEntries = MinimalEntries(contentTypes).ToList();
        spacedEntries.Add(("word/opaque.xml", Utf8("<opaque><a/> <b/></opaque>")));

        var compact = PackageManifestGenerator.Generate(BuildZip(
            compactEntries, CompressionLevel.NoCompression, DefaultTimestamp));
        var spaced = PackageManifestGenerator.Generate(BuildZip(
            spacedEntries, CompressionLevel.NoCompression, DefaultTimestamp));

        Assert.NotEqual(compact.Entries.Single(entry => entry.Uri == "/word/opaque.xml")
                .NormalizedXmlDigest,
            spaced.Entries.Single(entry => entry.Uri == "/word/opaque.xml")
                .NormalizedXmlDigest);
        Assert.NotEqual(compact.NormalizedSemanticDigest, spaced.NormalizedSemanticDigest);
    }

    [Fact]
    public void PM024_OpcAttributesMustBeUnqualifiedAndDefaultExtensionsAreSingleTokens()
    {
        const string contentTypes = """
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"
                   xmlns:fake="urn:fake">
              <Default fake:Extension="xml" fake:ContentType="application/xml"/>
              <Default Extension="" ContentType="application/xml"/>
              <Default Extension="." ContentType="application/xml"/>
              <Default Extension="path/xml" ContentType="application/xml"/>
              <Default Extension="path\xml" ContentType="application/xml"/>
              <Default Extension="control&#x7f;key" ContentType="application/xml"/>
              <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
            </Types>
            """;
        const string relationships = """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"
                           xmlns:fake="urn:fake">
              <Relationship fake:Id="rIdFake" fake:Type="urn:fake:type"
                            fake:Target="word/document.xml" fake:TargetMode="Internal"/>
            </Relationships>
            """;
        var entries = MinimalEntries(contentTypes).ToList();
        Replace(entries, "_rels/.rels", Utf8(relationships));

        var manifest = PackageManifestGenerator.Generate(BuildZip(
            entries, CompressionLevel.NoCompression, DefaultTimestamp));

        Assert.True(manifest.Findings.Count(finding =>
            finding.Code == "malformed_content_type") >= 2);
        Assert.Equal(4, manifest.Findings.Count(finding =>
            finding.Code == "invalid_content_type_extension"));
        Assert.Contains(manifest.Findings, finding => finding.Code == "malformed_relationship");
        Assert.Empty(manifest.Relationships);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PM025_DiagramRelationshipAttributesParticipateInClosure(bool strict)
    {
        var manifest = PackageManifestGenerator.Generate(BuildRichPackage(
            strict: strict,
            documentText: """
                <dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
                            r:dm="rIdDiagramData" r:lo="rIdDiagramLayout"
                            r:qs="rIdDiagramStyle" r:cs="rIdDiagramColors"/>
                """));

        var danglingIds = manifest.Findings
            .Where(finding => finding.Code == "dangling_relationship")
            .Select(finding => finding.Location?.RelationshipId)
            .ToHashSet(StringComparer.Ordinal);
        Assert.True(danglingIds.SetEquals(new HashSet<string>(StringComparer.Ordinal)
        {
            "rIdDiagramData",
            "rIdDiagramLayout",
            "rIdDiagramStyle",
            "rIdDiagramColors",
        }));
    }

    [Fact]
    public void PM026_WordFactsIgnoreSameLocalNameAttributesFromOtherNamespaces()
    {
        var package = BuildRichPackage(documentText: """
            <w:fldChar xmlns:fake="urn:fake" fake:fldCharType="begin"/>
            """);
        package = RewriteEntry(package, "word/footnotes.xml", xml => xml
            .Replace("xmlns:w=", "xmlns:fake=\"urn:fake\" xmlns:w=", StringComparison.Ordinal)
            .Replace("w:id=\"1\"", "fake:id=\"99\"", StringComparison.Ordinal));
        package = RewriteEntry(package, "word/commentsExtended.xml", xml => xml
            .Replace("xmlns:w15=", "xmlns:fake=\"urn:fake\" xmlns:w15=", StringComparison.Ordinal)
            .Replace("w15:done=", "fake:done=", StringComparison.Ordinal)
            .Replace("w15:paraIdParent=", "fake:paraIdParent=", StringComparison.Ordinal));

        var manifest = PackageManifestGenerator.Generate(package);

        Assert.Equal(1, manifest.Facts.FieldCount);
        Assert.Equal(0, manifest.Facts.FootnoteCount);
        Assert.Equal(0, manifest.Facts.Annotations.CommentReplies);
        Assert.Equal(0, manifest.Facts.Annotations.ResolvedComments);
    }

    [Fact]
    public void PM027_PackageAbsoluteRelationshipTargets_AreValidInternalPaths()
    {
        var package = RewriteEntry(
            BuildZip(MinimalEntries(), CompressionLevel.NoCompression, DefaultTimestamp),
            "_rels/.rels",
            xml => xml.Replace("Target=\"word/document.xml\"",
                "Target=\"/word/document.xml\"", StringComparison.Ordinal));

        var manifest = PackageManifestGenerator.Generate(package);

        Assert.DoesNotContain(manifest.Findings,
            finding => finding.Code == "invalid_relationship_target");
        var relationship = Assert.Single(manifest.Relationships,
            item => item.OwnerUri == "/" && item.Target == "/word/document.xml");
        Assert.Equal("/word/document.xml", relationship.ResolvedTargetUri);
        Assert.True(relationship.IsTargetPresent);
    }

    [Fact]
    public void PM028_DirectoryOnlyZipEntries_AreWarningsAndDoNotInvalidateThePackage()
    {
        var entries = MinimalEntries().ToList();
        entries.Insert(0, ("word/", Array.Empty<byte>()));
        entries.Insert(1, ("_rels/", Array.Empty<byte>()));

        var manifest = PackageManifestGenerator.Generate(
            BuildZip(entries, CompressionLevel.NoCompression, DefaultTimestamp));

        Assert.True(manifest.IsValid, string.Join("\n", manifest.Findings.Select(f => f.Code)));
        Assert.DoesNotContain(manifest.Findings, finding => finding.Code == "unsafe_entry_path");
        var directories = manifest.Findings.Where(f => f.Code == "directory_entry").ToList();
        Assert.Equal(2, directories.Count);
        Assert.All(directories,
            finding => Assert.Equal(VerificationFindingSeverity.Warning, finding.Severity));
        Assert.Contains(manifest.Entries, entry => entry.Uri == "/word/");
    }

    [Fact]
    public void PM029_DirectoryEntries_DoNotPerturbContentIdentities()
    {
        var withDirectories = MinimalEntries().ToList();
        withDirectories.Insert(0, ("word/", Array.Empty<byte>()));

        var plain = PackageManifestGenerator.Generate(
            BuildZip(MinimalEntries(), CompressionLevel.NoCompression, DefaultTimestamp));
        var padded = PackageManifestGenerator.Generate(
            BuildZip(withDirectories, CompressionLevel.NoCompression, DefaultTimestamp));

        Assert.Equal(plain.OrderedOpcContentDigest, padded.OrderedOpcContentDigest);
        Assert.Equal(plain.NormalizedSemanticDigest, padded.NormalizedSemanticDigest);
    }

    [Fact]
    public void PM030_EntryCountTruncation_SuppressesContentIdentities()
    {
        var options = new PackageManifestOptions { MaxEntryCount = 5 };

        var first = PackageManifestGenerator.Generate(TruncationPackage(0x01), options);
        var second = PackageManifestGenerator.Generate(TruncationPackage(0xfe), options);

        Assert.Contains(first.Findings, finding => finding.Code == "entry_count_limit_exceeded");
        Assert.NotEqual(first.RawPackageBytesDigest, second.RawPackageBytesDigest);
        Assert.Null(first.OrderedOpcContentDigest);
        Assert.Null(first.NormalizedSemanticDigest);
    }

    [Fact]
    public void PM031_DeclaredExpansionTotal_CoversEntriesBeyondTheInspectionLimit()
    {
        var entries = MinimalEntries().ToList();
        for (var index = 0; index < 20; index++)
            entries.Add(($"word/blob{index:D2}.bin", new byte[4096]));

        var manifest = PackageManifestGenerator.Generate(
            BuildZip(entries, CompressionLevel.NoCompression, DefaultTimestamp),
            new PackageManifestOptions { MaxEntryCount = 5, MaxTotalUncompressedBytes = 20_000 });

        Assert.Contains(manifest.Findings,
            finding => finding.Code == "total_expansion_limit_exceeded");
    }

    [Fact]
    public void PM032_OversizeContentTypes_ReportsTheXmlLimitOnce()
    {
        var entries = MinimalEntries().ToList();
        entries[0] = ("[Content_Types].xml", Utf8(OversizeContentTypes()));

        var manifest = PackageManifestGenerator.Generate(
            BuildZip(entries, CompressionLevel.NoCompression, DefaultTimestamp),
            new PackageManifestOptions { MaxXmlPartBytes = 1500 });

        Assert.Single(manifest.Findings, finding =>
            finding.Code == "xml_size_limit_exceeded"
            && finding.Location?.EntryUri == "/[Content_Types].xml");
    }

    [Fact]
    public void PM033_UnreadableContentTypes_ReportOnceInsteadOfPerEntry()
    {
        var package = RewriteEntry(
            BuildZip(MinimalEntries(), CompressionLevel.NoCompression, DefaultTimestamp),
            "[Content_Types].xml", _ => "<not-types/>");

        var manifest = PackageManifestGenerator.Generate(package);

        Assert.Contains(manifest.Findings, finding => finding.Code == "malformed_content_types");
        Assert.Contains(manifest.Findings, finding => finding.Code == "content_types_unreadable");
        Assert.DoesNotContain(manifest.Findings, finding => finding.Code == "missing_content_type");
        Assert.Equal("unresolved",
            manifest.Entries.Single(entry => entry.Uri == "/word/document.xml").ContentTypeSource);
    }

    [Fact]
    public void PM034_UnreadableRelationshipPart_IsReportedInsteadOfFakeDanglingIds()
    {
        var entries = MinimalEntries().ToList();
        entries[2] = ("word/document.xml", Utf8("""
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <w:body><w:p><w:hyperlink r:id="rIdReal"><w:r><w:t>x</w:t></w:r></w:hyperlink></w:p></w:body>
            </w:document>
            """));
        entries.Add(("word/_rels/document.xml.rels", Utf8(
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
            + "<!--" + new string('p', 4000) + "-->"
            + "<Relationship Id=\"rIdReal\" Type=\"urn:t\" Target=\"https://example.test\" TargetMode=\"External\"/>"
            + "</Relationships>")));

        var manifest = PackageManifestGenerator.Generate(
            BuildZip(entries, CompressionLevel.NoCompression, DefaultTimestamp),
            new PackageManifestOptions { MaxXmlPartBytes = 2000 });

        Assert.DoesNotContain(manifest.Findings, finding => finding.Code == "dangling_relationship");
        Assert.Contains(manifest.Findings, finding =>
            finding.Code == "relationship_part_unreadable"
            && finding.Location?.EntryUri == "/word/_rels/document.xml.rels");
    }

    [Fact]
    public void PM035_UnparsableXmlPart_KeepsAPackageSemanticIdentity()
    {
        var left = PackageManifestGenerator.Generate(MalformedXmlPackage("<open>"));
        var right = PackageManifestGenerator.Generate(MalformedXmlPackage("<other>"));

        Assert.Contains(left.Findings, finding => finding.Code == "malformed_xml");
        Assert.NotNull(left.NormalizedSemanticDigest);
        Assert.NotEqual(left.NormalizedSemanticDigest, right.NormalizedSemanticDigest);
    }

    [Fact]
    public void PM036_UnparsableXmlPart_StillIgnoresSerializationOfTheReadableParts()
    {
        var compact = MalformedXmlPackage("<open>");
        var pretty = RewriteEntry(compact, "word/document.xml",
            xml => xml.Replace("><", ">\n  <", StringComparison.Ordinal));

        var left = PackageManifestGenerator.Generate(compact).NormalizedSemanticDigest;
        var right = PackageManifestGenerator.Generate(pretty).NormalizedSemanticDigest;

        Assert.NotNull(left);
        Assert.Equal(left, right);
    }

    [Fact]
    public void PM037_BlankSessionManifest_IsValid()
    {
        var handle = DocxSessionOps.OpenSession(DocxSessionOps.CreateBlankDocx(), settings: null);
        try
        {
            using var parsed = JsonDocument.Parse(VerificationOps.GetPackageManifest(handle));
            var codes = parsed.RootElement.GetProperty("findings").EnumerateArray()
                .Select(finding => finding.GetProperty("code").GetString())
                .ToList();
            Assert.True(parsed.RootElement.GetProperty("isValid").GetBoolean(),
                string.Join(",", codes));
        }
        finally
        {
            DocxSessionOps.CloseSession(handle);
        }
    }

    [Fact]
    public void PM038_EveryCommittedDocxFixture_ProducesAValidManifest()
    {
        // The manifest is a verification artifact: a real Word-authored package that Word opens
        // must not be reported invalid. Fixtures listed here are genuinely malformed. The listing
        // pins the exact error codes rather than skipping the file, so a fixture that gets
        // corrected — or one that starts failing differently — fails this test instead of hiding.
        var known = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            // Its Override declares /word/afchunk2.dat as wordprocessingml.document.main+xml
            // while the payload is a ZIP, so the part cannot be parsed as the XML it claims.
            ["CA009-altChunk.docx"] = "malformed_xml",
        };
        var directory = new DirectoryInfo("../../../../TestFiles/");
        Assert.True(directory.Exists);

        var unexpected = new List<string>();
        foreach (var file in directory.EnumerateFiles("*.docx", SearchOption.AllDirectories)
                     .OrderBy(file => file.Name, StringComparer.Ordinal))
        {
            var manifest = PackageManifestGenerator.Generate(File.ReadAllBytes(file.FullName));
            var codes = string.Join(",", manifest.Findings
                .Where(finding => finding.Severity == VerificationFindingSeverity.Error)
                .Select(finding => finding.Code).Distinct().OrderBy(code => code, StringComparer.Ordinal));
            var expected = known.GetValueOrDefault(file.Name, string.Empty);
            if (codes != expected)
                unexpected.Add($"{file.Name}: expected [{expected}] got [{codes}]");
        }

        Assert.Empty(unexpected);
    }

    [Fact]
    public void PM040_XmlSkippedByASizeLimit_LeavesTheSemanticIdentityUnavailable()
    {
        // A part skipped for budget reasons could still have normalized under a larger budget.
        // Substituting its raw bytes would make the package identity a function of the caller's
        // options, so the digest is unavailable rather than different.
        var entries = MinimalEntries().ToList();
        entries.Add(("word/big.xml", Utf8("<big>" + new string('q', 3000) + "</big>")));
        var package = BuildZip(entries, CompressionLevel.NoCompression, DefaultTimestamp);

        var limited = PackageManifestGenerator.Generate(package,
            new PackageManifestOptions { MaxXmlPartBytes = 1500 });
        var full = PackageManifestGenerator.Generate(package);

        Assert.Contains(limited.Findings, finding => finding.Code == "xml_size_limit_exceeded");
        Assert.Null(limited.NormalizedSemanticDigest);
        Assert.NotNull(full.NormalizedSemanticDigest);
    }

    [Fact]
    public void PM041_TotalExpansionBreach_DoesNotBlameEachRelationshipPart()
    {
        var entries = MinimalEntries().ToList();
        entries.Add(("word/_rels/document.xml.rels", Utf8(
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"/>")));

        var manifest = PackageManifestGenerator.Generate(
            BuildZip(entries, CompressionLevel.NoCompression, DefaultTimestamp),
            new PackageManifestOptions { MaxTotalUncompressedBytes = 32 });

        Assert.Contains(manifest.Findings,
            finding => finding.Code == "total_expansion_limit_exceeded");
        Assert.DoesNotContain(manifest.Findings,
            finding => finding.Code == "relationship_part_unreadable");
        Assert.DoesNotContain(manifest.Findings,
            finding => finding.Code == "dangling_relationship");
    }

    [Fact]
    public void PM042_ContentTypeDeclarations_PreserveTheirDeclaredExtensionSpelling()
    {
        var entries = MinimalEntries("""
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="XML" ContentType="application/xml"/>
              <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
            </Types>
            """).ToList();

        var manifest = PackageManifestGenerator.Generate(
            BuildZip(entries, CompressionLevel.NoCompression, DefaultTimestamp));

        Assert.Contains(manifest.ContentTypes,
            declaration => declaration.Kind == "default" && declaration.Key == "XML");
        Assert.Equal("application/vnd.openxmlformats-package.relationships+xml",
            manifest.Entries.Single(entry => entry.Uri == "/_rels/.rels").ContentType);
    }

    [Fact]
    public void PM039_RelationshipReferenceOnTheRootElement_IsCheckedForClosure()
    {
        var entries = MinimalEntries().ToList();
        entries[2] = ("word/document.xml", Utf8("""
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                        r:id="rIdOnRoot"><w:body><w:p/></w:body></w:document>
            """));

        var manifest = PackageManifestGenerator.Generate(
            BuildZip(entries, CompressionLevel.NoCompression, DefaultTimestamp));

        Assert.Contains(manifest.Findings, finding =>
            finding.Code == "dangling_relationship"
            && finding.Location?.RelationshipId == "rIdOnRoot");
    }

    private static byte[] TruncationPackage(byte tail)
    {
        var entries = MinimalEntries().ToList();
        for (var index = 0; index < 8; index++)
            entries.Add(($"word/extra{index}.bin", new[] { (byte)index }));
        entries.Add(("word/zzz-payload.bin", new[] { tail, tail, tail }));
        return BuildZip(entries, CompressionLevel.NoCompression, DefaultTimestamp);
    }

    private static byte[] MalformedXmlPackage(string payload)
    {
        var entries = MinimalEntries().ToList();
        entries.Add(("word/broken.xml", Utf8(payload)));
        return BuildZip(entries, CompressionLevel.NoCompression, DefaultTimestamp);
    }

    private static string OversizeContentTypes() => """
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="xml" ContentType="application/xml"/>
          <Default Extension="bin" ContentType="application/octet-stream"/>
          <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
        """ + "<!--" + new string('q', 3000) + "--></Types>";

    private static readonly DateTimeOffset DefaultTimestamp =
        new(2024, 1, 1, 0, 0, 0, TimeSpan.Zero);

    private static byte[] BuildRichPackage(
        bool strict = false,
        bool reverseEntries = false,
        CompressionLevel compression = CompressionLevel.Optimal,
        DateTimeOffset? timestamp = null,
        string? customXml = null,
        string? documentText = null)
    {
        var w = strict
            ? "http://purl.oclc.org/ooxml/wordprocessingml/main"
            : "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var r = strict
            ? "http://purl.oclc.org/ooxml/officeDocument/relationships"
            : "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var entries = new List<(string Name, byte[] Bytes)>
        {
            ("[Content_Types].xml", Utf8(ContentTypes())),
            ("_rels/.rels", Utf8($"""
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdMain" Type="{r}/officeDocument" Target="word/document.xml"/>
                  <Relationship Id="rIdCore" Type="{r}/metadata/core-properties" Target="docProps/core.xml"/>
                  <Relationship Id="rIdApp" Type="{r}/extended-properties" Target="docProps/app.xml"/>
                  <Relationship Id="rIdCustomProps" Type="{r}/custom-properties" Target="docProps/custom.xml"/>
                </Relationships>
                """)),
            ("word/document.xml", Utf8($"""
                <w:document xmlns:w="{w}" xmlns:r="{r}"><w:body>
                  <w:p>{documentText ?? "<w:r><w:t>Hello</w:t></w:r>"}<w:hyperlink r:id="rIdExternal"><w:r><w:t>link</w:t></w:r></w:hyperlink><w:r><w:drawing><x/></w:drawing></w:r><w:fldSimple w:instr="PAGE"/><w:altChunk r:id="rIdChunk"/></w:p>
                  <w:p><w:ins><w:r><w:t>i</w:t></w:r></w:ins><w:del><w:r><w:delText>d</w:delText></w:r></w:del><w:moveFrom><w:r><w:t>m1</w:t></w:r></w:moveFrom><w:moveTo><w:r><w:t>m2</w:t></w:r></w:moveTo><w:pPrChange/></w:p>
                  <w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
                  <w:sectPr><w:headerReference r:id="rIdHeader"/><w:footerReference r:id="rIdFooter"/></w:sectPr>
                </w:body></w:document>
                """)),
            ("word/_rels/document.xml.rels", Utf8($"""
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rIdHeader" Type="{r}/header" Target="header1.xml"/>
                  <Relationship Id="rIdFooter" Type="{r}/footer" Target="footer1.xml"/>
                  <Relationship Id="rIdFootnotes" Type="{r}/footnotes" Target="footnotes.xml"/>
                  <Relationship Id="rIdEndnotes" Type="{r}/endnotes" Target="endnotes.xml"/>
                  <Relationship Id="rIdComments" Type="{r}/comments" Target="comments.xml"/>
                  <Relationship Id="rIdCommentsEx" Type="{r}/commentsExtended" Target="commentsExtended.xml"/>
                  <Relationship Id="rIdPeople" Type="{r}/people" Target="people.xml"/>
                  <Relationship Id="rIdStyles" Type="{r}/styles" Target="styles.xml"/>
                  <Relationship Id="rIdNumbering" Type="{r}/numbering" Target="numbering.xml"/>
                  <Relationship Id="rIdTheme" Type="{r}/theme" Target="theme/theme1.xml"/>
                  <Relationship Id="rIdImage" Type="{r}/image" Target="media/image1.png"/>
                  <Relationship Id="rIdCustom" Type="{r}/customXml" Target="../customXml/item2.xml"/>
                  <Relationship Id="rIdChunk" Type="{r}/aFChunk" Target="afchunk.html"/>
                  <Relationship Id="rIdExternal" Type="{r}/hyperlink" Target="https://example.test" TargetMode="External"/>
                </Relationships>
                """)),
            ("word/header1.xml", Utf8($"<w:hdr xmlns:w=\"{w}\"><w:p><w:r><w:t>header</w:t></w:r></w:p></w:hdr>")),
            ("word/footer1.xml", Utf8($"<w:ftr xmlns:w=\"{w}\"><w:p><w:r><w:t>footer</w:t></w:r></w:p></w:ftr>")),
            ("word/footnotes.xml", Utf8($"<w:footnotes xmlns:w=\"{w}\"><w:footnote w:id=\"-1\"><w:p/></w:footnote><w:footnote w:id=\"1\"><w:p><w:r><w:t>foot</w:t></w:r></w:p></w:footnote></w:footnotes>")),
            ("word/endnotes.xml", Utf8($"<w:endnotes xmlns:w=\"{w}\"><w:endnote w:id=\"0\"><w:p/></w:endnote><w:endnote w:id=\"1\"><w:p><w:r><w:t>end</w:t></w:r></w:p></w:endnote></w:endnotes>")),
            ("word/comments.xml", Utf8($"<w:comments xmlns:w=\"{w}\"><w:comment w:id=\"1\"><w:p/></w:comment><w:comment w:id=\"2\"><w:p/></w:comment></w:comments>")),
            ("word/commentsExtended.xml", Utf8("<w15:commentsEx xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\"><w15:commentEx w15:paraId=\"1\" w15:done=\"1\"/><w15:commentEx w15:paraId=\"2\" w15:paraIdParent=\"1\" w15:done=\"0\"/></w15:commentsEx>")),
            ("word/people.xml", Utf8("<w15:people xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\"><w15:person w15:author=\"A\"/></w15:people>")),
            ("word/styles.xml", Utf8($"<w:styles xmlns:w=\"{w}\"><w:style w:styleId=\"Normal\"/><w:style w:styleId=\"Heading1\"/></w:styles>")),
            ("word/numbering.xml", Utf8($"<w:numbering xmlns:w=\"{w}\"><w:abstractNum w:abstractNumId=\"1\"/><w:num w:numId=\"1\"/></w:numbering>")),
            ("word/theme/theme1.xml", Utf8("<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Theme\"/>")),
            ("word/media/image1.png", new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }),
            ("word/afchunk.html", Utf8("<p>chunk</p>")),
            ("word/data/payload.weird", new byte[] { 4, 5, 6 }),
            ("customXml/item1.xml", Utf8("<annotations xmlns=\"http://docxodus.dev/annotations/v1\"><annotation id=\"a1\"/></annotations>")),
            ("customXml/item2.xml", Utf8(customXml ?? "<data xmlns=\"urn:opaque\"><value>opaque</value></data>")),
            ("customXml/itemProps1.xml", Utf8("<ds:datastoreItem xmlns:ds=\"http://schemas.openxmlformats.org/officeDocument/2006/customXml\" ds:itemID=\"{1}\"/>")),
            ("docProps/core.xml", Utf8("<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\"/>")),
            ("docProps/app.xml", Utf8("<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\"/>")),
            ("docProps/custom.xml", Utf8("<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/custom-properties\"/>")),
        };
        if (reverseEntries)
            entries.Reverse();
        return BuildZip(entries, compression, timestamp ?? DefaultTimestamp);
    }

    private static string ContentTypes() => """
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Default Extension="xml" ContentType="application/xml"/>
          <Default Extension="png" ContentType="image/png"/>
          <Default Extension="html" ContentType="text/html"/>
          <Default Extension="weird" ContentType="application/x-docxodus-test"/>
          <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
          <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
          <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
          <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
          <Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>
          <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
          <Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.ms-word.commentsExt+xml"/>
          <Override PartName="/word/people.xml" ContentType="application/vnd.ms-word.people+xml"/>
          <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
          <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
          <Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
          <Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>
          <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
          <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
          <Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>
        </Types>
        """;

    private static IEnumerable<(string Name, byte[] Bytes)> MinimalEntries(string? contentTypes = null)
    {
        yield return ("[Content_Types].xml", Utf8(contentTypes ?? """
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="xml" ContentType="application/xml"/>
              <Default Extension="bin" ContentType="application/octet-stream"/>
              <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
            </Types>
            """));
        yield return ("_rels/.rels", Utf8("""
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
            </Relationships>
            """));
        yield return ("word/document.xml", Utf8("""
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>
            """));
    }

    private static byte[] BuildZip(
        IEnumerable<(string Name, byte[] Bytes)> entries,
        CompressionLevel compression,
        DateTimeOffset timestamp)
    {
        using var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var (name, bytes) in entries)
            {
                var entry = archive.CreateEntry(name, compression);
                entry.LastWriteTime = timestamp;
                using var output = entry.Open();
                output.Write(bytes);
            }
        }
        return stream.ToArray();
    }

    private static byte[] RewriteEntry(byte[] package, string name, Func<string, string> rewrite)
    {
        var entries = new List<(string Name, byte[] Bytes)>();
        using (var input = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read))
        {
            foreach (var entry in input.Entries)
            {
                using var stream = entry.Open();
                using var copy = new MemoryStream();
                stream.CopyTo(copy);
                var bytes = copy.ToArray();
                if (entry.FullName == name)
                    bytes = Utf8(rewrite(Encoding.UTF8.GetString(bytes)));
                entries.Add((entry.FullName, bytes));
            }
        }
        return BuildZip(entries, CompressionLevel.Optimal, DefaultTimestamp);
    }

    private static byte[] MarkFirstEntryEncrypted(byte[] bytes)
    {
        var copy = bytes.ToArray();
        var local = FindSignature(copy, 0x04034b50);
        var central = FindSignature(copy, 0x02014b50);
        Assert.True(local >= 0 && central >= 0);
        BinaryPrimitives.WriteUInt16LittleEndian(copy.AsSpan(local + 6, 2),
            (ushort)(BinaryPrimitives.ReadUInt16LittleEndian(copy.AsSpan(local + 6, 2)) | 1));
        BinaryPrimitives.WriteUInt16LittleEndian(copy.AsSpan(central + 8, 2),
            (ushort)(BinaryPrimitives.ReadUInt16LittleEndian(copy.AsSpan(central + 8, 2)) | 1));
        return copy;
    }

    private static byte[] PromoteToZip64Directory(byte[] bytes)
    {
        const uint endOfCentralDirectorySignature = 0x06054b50;
        var end = FindSignature(bytes, endOfCentralDirectorySignature);
        Assert.True(end >= 0);
        var entryCount = BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(end + 10, 2));
        var centralSize = BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(end + 12, 4));
        var centralOffset = BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(end + 16, 4));
        var zip64End = new byte[56];
        BinaryPrimitives.WriteUInt32LittleEndian(zip64End.AsSpan(0, 4), 0x06064b50);
        BinaryPrimitives.WriteUInt64LittleEndian(zip64End.AsSpan(4, 8), 44);
        BinaryPrimitives.WriteUInt16LittleEndian(zip64End.AsSpan(12, 2), 45);
        BinaryPrimitives.WriteUInt16LittleEndian(zip64End.AsSpan(14, 2), 45);
        BinaryPrimitives.WriteUInt64LittleEndian(zip64End.AsSpan(24, 8), entryCount);
        BinaryPrimitives.WriteUInt64LittleEndian(zip64End.AsSpan(32, 8), entryCount);
        BinaryPrimitives.WriteUInt64LittleEndian(zip64End.AsSpan(40, 8), centralSize);
        BinaryPrimitives.WriteUInt64LittleEndian(zip64End.AsSpan(48, 8), centralOffset);

        var locator = new byte[20];
        BinaryPrimitives.WriteUInt32LittleEndian(locator.AsSpan(0, 4), 0x07064b50);
        BinaryPrimitives.WriteUInt64LittleEndian(locator.AsSpan(8, 8), (ulong)end);
        BinaryPrimitives.WriteUInt32LittleEndian(locator.AsSpan(16, 4), 1);

        var classicEnd = bytes.AsSpan(end, 22).ToArray();
        BinaryPrimitives.WriteUInt16LittleEndian(classicEnd.AsSpan(8, 2), ushort.MaxValue);
        BinaryPrimitives.WriteUInt16LittleEndian(classicEnd.AsSpan(10, 2), ushort.MaxValue);
        BinaryPrimitives.WriteUInt32LittleEndian(classicEnd.AsSpan(12, 4), uint.MaxValue);
        BinaryPrimitives.WriteUInt32LittleEndian(classicEnd.AsSpan(16, 4), uint.MaxValue);

        using var output = new MemoryStream();
        output.Write(bytes, 0, end);
        output.Write(zip64End);
        output.Write(locator);
        output.Write(classicEnd);
        return output.ToArray();
    }

    private static byte[] RewriteCentralUncompressedSizes(
        byte[] bytes,
        IReadOnlyDictionary<string, uint> sizes)
    {
        const uint endOfCentralDirectorySignature = 0x06054b50;
        const uint centralDirectorySignature = 0x02014b50;
        var copy = bytes.ToArray();
        var end = FindSignature(copy, endOfCentralDirectorySignature);
        Assert.True(end >= 0);
        var count = BinaryPrimitives.ReadUInt16LittleEndian(copy.AsSpan(end + 10, 2));
        var position = checked((int)BinaryPrimitives.ReadUInt32LittleEndian(
            copy.AsSpan(end + 16, 4)));
        for (var index = 0; index < count; index++)
        {
            Assert.Equal(centralDirectorySignature,
                BinaryPrimitives.ReadUInt32LittleEndian(copy.AsSpan(position, 4)));
            var nameLength = BinaryPrimitives.ReadUInt16LittleEndian(
                copy.AsSpan(position + 28, 2));
            var extraLength = BinaryPrimitives.ReadUInt16LittleEndian(
                copy.AsSpan(position + 30, 2));
            var commentLength = BinaryPrimitives.ReadUInt16LittleEndian(
                copy.AsSpan(position + 32, 2));
            var name = Encoding.UTF8.GetString(copy, position + 46, nameLength);
            if (sizes.TryGetValue(name, out var size))
                BinaryPrimitives.WriteUInt32LittleEndian(copy.AsSpan(position + 24, 4), size);
            position += 46 + nameLength + extraLength + commentLength;
        }
        return copy;
    }

    private static int FindSignature(byte[] bytes, uint signature)
    {
        for (var index = 0; index <= bytes.Length - 4; index++)
            if (BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(index, 4)) == signature)
                return index;
        return -1;
    }

    private static void Replace(
        List<(string Name, byte[] Bytes)> entries,
        string name,
        byte[] replacement)
    {
        var index = entries.FindIndex(entry => entry.Name == name);
        Assert.True(index >= 0);
        entries[index] = (name, replacement);
    }

    private static byte[] Utf8(string value) => Encoding.UTF8.GetBytes(value);
}
