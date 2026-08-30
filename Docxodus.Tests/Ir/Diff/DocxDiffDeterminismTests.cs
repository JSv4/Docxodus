// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.IO.Packaging;
using System.Linq;
using System.Security.Cryptography;
using System.Xml.Linq;
using Docxodus;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// <see cref="Docxodus.Ir.Diff.IrDiffSettings.Deterministic"/> promises that two comparisons of the same
/// inputs are byte-identical. That held for text-only documents and failed for every redline that created or
/// imported a part: imported parts were named <c>P</c> + a fresh <c>Guid</c>, the relationships pointing at
/// them got <c>R</c> + a fresh <c>Guid</c>, and parts added through the SDK's <c>AddNewPart&lt;T&gt;()</c>
/// took the SDK's own <c>"R"</c> + sixteen random hex characters. The churn reached <c>document.xml</c>, the
/// <c>_rels</c> and the <c>[Content_Types].xml</c> overrides, so the same comparison produced a different
/// artifact hash every time — defeating content-addressed storage, caching, signing and byte-level
/// regression testing alike.
/// </summary>
public class DocxDiffDeterminismTests
{
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    /// <summary>A redline that imports a raster image: the media part, its relationship, and the
    /// <c>a:blip/@r:embed</c> that names it all used to move on every run.</summary>
    [Fact]
    public void Compare_TwiceOverImportedImage_IsByteIdentical()
    {
        AssertCompareIsReproducible("WC/WC013-Image-Before.docx", "WC/WC013-Image-After.docx");
    }

    /// <summary>The worst case in the corpus: one SmartArt comparison churned nine
    /// <c>word/diagrams/P*.xml</c> parts plus their <c>_rels</c> and the matching content-type overrides.</summary>
    [Fact]
    public void Compare_TwiceOverImportedSmartArt_IsByteIdentical()
    {
        var output = AssertCompareIsReproducible(
            "WC/WC014-SmartArt-With-Image-Before.docx", "WC/WC014-SmartArt-With-Image-After.docx");

        // Guard the premise: this fixture really does exercise the multi-part diagram import.
        Assert.True(output.Keys.Count(k => k.StartsWith("word/diagrams/", StringComparison.Ordinal)) > 1);
        Assert.Contains(output.Keys, k => k.StartsWith("word/media/", StringComparison.Ordinal));
    }

    /// <summary>Parts the renderer CREATES rather than imports — here a numbering part for a right-only
    /// numbered list — are the second source of churn: <c>AddNewPart&lt;T&gt;()</c> with no explicit id
    /// takes a random one from the SDK. No media is involved, so this fails independently of the import path.
    /// </summary>
    [Fact]
    public void Compare_TwiceOverRendererCreatedParts_IsByteIdentical()
    {
        var output = AssertCompareIsReproducible("Blank-wml.docx", "CA/CA003-Numbered-List.docx");
        Assert.Contains("word/numbering.xml", output.Keys);
    }

    /// <summary>
    /// The naming rule itself, not just its stability: an imported media part is named for a hash of its own
    /// bytes, and those bytes are the source part's verbatim. Pinning both together is what separates a
    /// reproducible import from one that reproducibly writes the wrong thing.
    /// </summary>
    [Fact]
    public void Compare_ImportedMediaPart_IsNamedByTheContentAddressOfItsUnchangedBytes()
    {
        var right = TestFile("WC/WC013-Image-After.docx");
        var sourceMedia = ZipEntries(File.ReadAllBytes(right))
            .Where(e => e.Key.StartsWith("word/media/", StringComparison.Ordinal))
            .Select(e => e.Value)
            .ToList();
        Assert.NotEmpty(sourceMedia);

        var output = ZipEntries(DocxDiff.Compare(
            new WmlDocument(TestFile("WC/WC013-Image-Before.docx")), new WmlDocument(right)).DocumentByteArray);

        var imported = output
            .Where(e => e.Key.StartsWith("word/media/P", StringComparison.Ordinal))
            .ToList();
        Assert.NotEmpty(imported);

        foreach (var part in imported)
        {
            // The bytes survived the copy untouched...
            Assert.Contains(sourceMedia, b => b.SequenceEqual(part.Value));
            // ...and the name is exactly the content address of those bytes, so two runs — and two
            // machines — cannot disagree about it.
            var name = Path.GetFileNameWithoutExtension(part.Key);
            Assert.Equal(
                Convert.ToHexStringLower(SHA256.HashData(part.Value))[..32],
                name.TrimStart('P')[..32]);
        }
    }

    /// <summary>
    /// Every XML part of a redline must still parse. The deterministic relationship ids are far SHORTER than
    /// the 33-character <c>R</c>+Guid they replaced, which exposed a latent bug in the import fixup: it
    /// rewrote a copied part through <c>GetStream()</c> (OpenOrCreate, no truncation), so a rewrite that
    /// shrank the part left the tail of the original behind and the part stopped parsing. Byte-equality
    /// between two runs cannot see that — two runs corrupt a part identically.
    /// </summary>
    [Theory]
    [InlineData("WC/WC014-SmartArt-With-Image-Before.docx", "WC/WC014-SmartArt-With-Image-After.docx")]
    [InlineData("CU009-Chart-Embedded-Xlsx-01.docx", "CU010-Chart-Embedded-Xlsx-02.docx")]
    public void Compare_ImportedXmlParts_RemainWellFormedAfterIdRemapping(string leftName, string rightName)
    {
        var output = ZipEntries(DocxDiff.Compare(
            new WmlDocument(TestFile(leftName)), new WmlDocument(TestFile(rightName))).DocumentByteArray);

        // Imported parts are the ones the fixup rewrites, and only they carry the "P"<address> name.
        var imported = output
            .Where(e => e.Key.Contains("/P", StringComparison.Ordinal)
                        && (e.Key.EndsWith(".xml", StringComparison.Ordinal)
                            || e.Key.EndsWith(".rels", StringComparison.Ordinal)))
            .ToList();
        Assert.NotEmpty(imported);

        foreach (var part in imported)
        {
            using var stream = new MemoryStream(part.Value);
            var parsed = Record.Exception(() => XDocument.Load(stream));
            Assert.True(parsed is null, $"{part.Key} did not parse: {parsed?.Message}");
        }
    }

    /// <summary>
    /// Content addressing folds byte-identical sources onto one name, but the importer must still be able to
    /// keep two copies apart — a shared SmartArt data part is cloned per owner precisely so each clone can be
    /// rewired differently. The suffix probe covers that: same bytes in, two distinct destination parts out.
    /// </summary>
    [Fact]
    public void MoveRelatedParts_TwoSourcePartsWithIdenticalBytes_GetDistinctDestinationParts()
    {
        const string xmlRelationship = "urn:docxdiff:test/xml";
        XNamespace test = "urn:docxdiff:identical-bytes";

        using var sourceBytes = new MemoryStream();
        using var destinationBytes = new MemoryStream();
        using var sourcePackage = Package.Open(sourceBytes, FileMode.Create, FileAccess.ReadWrite);
        using var destinationPackage = Package.Open(destinationBytes, FileMode.Create, FileAccess.ReadWrite);

        var sourceRoot = sourcePackage.CreatePart(new Uri("/source/root.xml", UriKind.Relative), "application/xml");
        var twins = new[] { "/source/twin-a.xml", "/source/twin-b.xml" }
            .Select(uri => sourcePackage.CreatePart(new Uri(uri, UriKind.Relative), "application/xml"))
            .ToList();
        WritePackageXml(sourceRoot, $"<t:root xmlns:t=\"{test}\"/>");
        foreach (var twin in twins)
            WritePackageXml(twin, $"<t:twin xmlns:t=\"{test}\">identical</t:twin>");

        sourceRoot.CreateRelationship(twins[0].Uri, TargetMode.Internal, xmlRelationship, "rIdA");
        sourceRoot.CreateRelationship(twins[1].Uri, TargetMode.Internal, xmlRelationship, "rIdB");

        var destinationRoot = destinationPackage.CreatePart(
            new Uri("/destination/root.xml", UriKind.Relative), "application/xml");
        var carrier = new XElement(test + "root",
            new XAttribute(XNamespace.Xmlns + "r", R),
            new XElement(test + "a", new XAttribute(R + "id", "rIdA")),
            new XElement(test + "b", new XAttribute(R + "id", "rIdB")));

        WmlComparer.MoveRelatedPartsToDestination(sourceRoot, destinationRoot, carrier);

        var a = RelatedPart(destinationRoot, (string)carrier.Element(test + "a")!.Attribute(R + "id")!);
        var b = RelatedPart(destinationRoot, (string)carrier.Element(test + "b")!.Attribute(R + "id")!);
        Assert.NotEqual(a.Uri, b.Uri);
        // Both are addressed by the same content, so one carries the disambiguating suffix.
        Assert.EndsWith("-1.xml", b.Uri.ToString(), StringComparison.Ordinal);
        Assert.Equal("identical", ReadPackageXml(a).Root!.Value);
        Assert.Equal("identical", ReadPackageXml(b).Root!.Value);
    }

    /// <summary>Compare twice and assert the whole OPC package matches part for part and byte for byte;
    /// returns the output's entries so a caller can assert what the fixture actually exercised.</summary>
    private static Dictionary<string, byte[]> AssertCompareIsReproducible(string leftName, string rightName)
    {
        var left = new WmlDocument(TestFile(leftName));
        var right = new WmlDocument(TestFile(rightName));

        var first = ZipEntries(DocxDiff.Compare(left, right).DocumentByteArray);
        var second = ZipEntries(DocxDiff.Compare(left, right).DocumentByteArray);

        Assert.Equal(
            first.Keys.OrderBy(k => k, StringComparer.Ordinal),
            second.Keys.OrderBy(k => k, StringComparer.Ordinal));
        foreach (var entry in first)
            Assert.True(
                entry.Value.SequenceEqual(second[entry.Key]),
                $"{entry.Key} differs between two comparisons of the same inputs.");
        return first;
    }

    private static readonly DirectoryInfo SourceDir = new("../../../../TestFiles/");

    private static string TestFile(string relativePath) =>
        Path.Combine(SourceDir.FullName, relativePath);

    private static Dictionary<string, byte[]> ZipEntries(byte[] package)
    {
        var entries = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        using var stream = new MemoryStream(package);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        foreach (var entry in archive.Entries)
        {
            using var entryStream = entry.Open();
            using var buffer = new MemoryStream();
            entryStream.CopyTo(buffer);
            entries[entry.FullName] = buffer.ToArray();
        }
        return entries;
    }

    private static PackagePart RelatedPart(PackagePart owner, string relationshipId)
    {
        var relationship = owner.GetRelationship(relationshipId);
        return owner.Package.GetPart(PackUriHelper.ResolvePartUri(owner.Uri, relationship.TargetUri));
    }

    private static XDocument ReadPackageXml(PackagePart part)
    {
        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        return XDocument.Load(stream);
    }

    private static void WritePackageXml(PackagePart part, string xml)
    {
        using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
        using var writer = new StreamWriter(stream);
        writer.Write(xml);
    }
}
