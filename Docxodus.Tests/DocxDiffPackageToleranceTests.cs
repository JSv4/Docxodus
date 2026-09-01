#nullable enable
using System;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using Docxodus;
using Docxodus.Tests.Ir;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Word-tolerated broken packages: a relationship whose internal target part is missing (for
/// example a <c>/docProps/thumbnail.jpeg</c> declared in <c>_rels/.rels</c> but absent from the
/// ZIP) opens fine in Word, which silently drops the dangling reference on save. The Open XML
/// SDK instead throws from its eager part-tree load the first time <c>MainDocumentPart</c> is
/// touched. These tests pin that both comparison inputs and the emitted redline behave like
/// Word: the compare succeeds and the output carries no dangling relationship.
/// </summary>
public class DocxDiffPackageToleranceTests
{
    private const string ThumbnailRelType =
        "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail";

    private const string ImageRelType =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

    /// <summary>Adds a package-level relationship pointing at a part that does not exist.</summary>
    private static WmlDocument WithDanglingPackageRelationship(WmlDocument doc)
    {
        using var ms = new MemoryStream();
        ms.Write(doc.DocumentByteArray, 0, doc.DocumentByteArray.Length);
        using (var package = Package.Open(ms, FileMode.Open, FileAccess.ReadWrite))
        {
            package.CreateRelationship(
                new Uri("/docProps/thumbnail.jpeg", UriKind.Relative),
                TargetMode.Internal,
                ThumbnailRelType);
        }

        return new WmlDocument(doc.FileName, ms.ToArray());
    }

    /// <summary>Adds a main-part relationship pointing at a media part that does not exist.</summary>
    private static WmlDocument WithDanglingPartRelationship(WmlDocument doc)
    {
        using var ms = new MemoryStream();
        ms.Write(doc.DocumentByteArray, 0, doc.DocumentByteArray.Length);
        using (var package = Package.Open(ms, FileMode.Open, FileAccess.ReadWrite))
        {
            var mainPart = package.GetPart(new Uri("/word/document.xml", UriKind.Relative));
            mainPart.CreateRelationship(
                new Uri("media/missing-image.png", UriKind.Relative),
                TargetMode.Internal,
                ImageRelType);
        }

        return new WmlDocument(doc.FileName, ms.ToArray());
    }

    private static void AssertNoDanglingRelationships(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var package = Package.Open(ms, FileMode.Open, FileAccess.Read);
        foreach (var rel in package.GetRelationships().Where(r => r.TargetMode == TargetMode.Internal))
        {
            var target = PackUriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
            Assert.True(package.PartExists(target), $"package relationship {rel.Id} -> {target} dangles");
        }

        foreach (var part in package.GetParts().Where(p => !PackUriHelper.IsRelationshipPartUri(p.Uri)))
        {
            foreach (var rel in part.GetRelationships().Where(r => r.TargetMode == TargetMode.Internal))
            {
                var target = PackUriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
                Assert.True(package.PartExists(target), $"{part.Uri} relationship {rel.Id} -> {target} dangles");
            }
        }
    }

    [Fact]
    public void Compare_tolerates_dangling_package_relationship_on_right()
    {
        var left = IrTestDocuments.FromBodyXml("<w:p><w:r><w:t>Alpha shared body text.</w:t></w:r></w:p>");
        var right = WithDanglingPackageRelationship(
            IrTestDocuments.FromBodyXml("<w:p><w:r><w:t>Alpha revised body text.</w:t></w:r></w:p>"));

        var redline = DocxDiff.Compare(left, right);

        Assert.NotNull(redline);
        AssertNoDanglingRelationships(redline);
    }

    [Fact]
    public void Compare_tolerates_dangling_package_relationship_on_left()
    {
        var left = WithDanglingPackageRelationship(
            IrTestDocuments.FromBodyXml("<w:p><w:r><w:t>Alpha shared body text.</w:t></w:r></w:p>"));
        var right = IrTestDocuments.FromBodyXml("<w:p><w:r><w:t>Alpha revised body text.</w:t></w:r></w:p>");

        var redline = DocxDiff.Compare(left, right);

        Assert.NotNull(redline);
        AssertNoDanglingRelationships(redline);
    }

    [Fact]
    public void Compare_tolerates_dangling_part_relationship()
    {
        var left = IrTestDocuments.FromBodyXml("<w:p><w:r><w:t>Alpha shared body text.</w:t></w:r></w:p>");
        var right = WithDanglingPartRelationship(
            IrTestDocuments.FromBodyXml("<w:p><w:r><w:t>Alpha revised body text.</w:t></w:r></w:p>"));

        var redline = DocxDiff.Compare(left, right);

        Assert.NotNull(redline);
        AssertNoDanglingRelationships(redline);
    }
}
