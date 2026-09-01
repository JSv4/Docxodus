#nullable enable
using System;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using Docxodus;
using Docxodus.Tests.Ir;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// ECMA-376 Strict inputs (purl.oclc.org namespaces, strict relationship types). Word converts
/// them to Transitional when it compares — its redline output of two strict documents is fully
/// transitional — and every raw-XML consumer downstream of us expects transitional names. These
/// tests pin that behavior: comparing strict inputs succeeds and yields a transitional package
/// with real tracked-changes markup.
/// </summary>
public class DocxDiffStrictConformanceTests
{
    private const string TransitionalMain = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string StrictMain = "http://purl.oclc.org/ooxml/wordprocessingml/main";
    private const string TransitionalRelPrefix = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
    private const string StrictRelPrefix = "http://purl.oclc.org/ooxml/officeDocument/relationships/";

    /// <summary>
    /// Rewrites a transitional test document into a minimal ECMA-376 Strict package: the
    /// wordprocessingml namespace flips to the strict URI in every XML part and the
    /// officeDocument-family relationship types flip to their strict (purl.oclc.org) forms.
    /// </summary>
    private static WmlDocument AsStrict(WmlDocument doc)
    {
        using var ms = new MemoryStream();
        ms.Write(doc.DocumentByteArray, 0, doc.DocumentByteArray.Length);
        using (var package = Package.Open(ms, FileMode.Open, FileAccess.ReadWrite))
        {
            var packageRels = package.GetRelationships()
                .Where(r => r.RelationshipType.StartsWith(TransitionalRelPrefix, StringComparison.Ordinal))
                .Select(r => (r.Id, r.TargetUri, r.RelationshipType))
                .ToList();
            foreach (var (id, target, relType) in packageRels)
            {
                package.DeleteRelationship(id);
                package.CreateRelationship(
                    target, TargetMode.Internal,
                    StrictRelPrefix + relType.Substring(TransitionalRelPrefix.Length), id);
            }

            foreach (var part in package.GetParts().ToList())
            {
                if (PackUriHelper.IsRelationshipPartUri(part.Uri))
                    continue;
                if (part.ContentType.EndsWith("+xml", StringComparison.Ordinal))
                {
                    string xml;
                    using (var reader = new StreamReader(part.GetStream(FileMode.Open, FileAccess.Read)))
                        xml = reader.ReadToEnd();
                    xml = xml.Replace(TransitionalMain, StrictMain);
                    using var writer = new StreamWriter(part.GetStream(FileMode.Create, FileAccess.Write), Encoding.UTF8);
                    writer.Write(xml);
                }

                var partRels = part.GetRelationships()
                    .Where(r => r.RelationshipType.StartsWith(TransitionalRelPrefix, StringComparison.Ordinal))
                    .Select(r => (r.Id, r.TargetUri, r.RelationshipType))
                    .ToList();
                foreach (var (id, target, relType) in partRels)
                {
                    part.DeleteRelationship(id);
                    part.CreateRelationship(
                        target, TargetMode.Internal,
                        StrictRelPrefix + relType.Substring(TransitionalRelPrefix.Length), id);
                }
            }
        }

        return new WmlDocument(doc.FileName, ms.ToArray());
    }

    private static (string MainXml, string OfficeDocumentRelType) ReadMain(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var package = Package.Open(ms, FileMode.Open, FileAccess.Read);
        var rel = package.GetRelationships()
            .Single(r => r.RelationshipType.EndsWith("officeDocument", StringComparison.Ordinal)
                      && r.RelationshipType.Contains("relationships"));
        var part = package.GetPart(PackUriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri));
        using var reader = new StreamReader(part.GetStream(FileMode.Open, FileAccess.Read));
        return (reader.ReadToEnd(), rel.RelationshipType);
    }

    [Fact]
    public void Compare_of_strict_inputs_yields_a_transitional_redline()
    {
        var left = AsStrict(IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>Alpha shared body text.</w:t></w:r></w:p>"));
        var right = AsStrict(IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>Alpha revised body text.</w:t></w:r></w:p>"));

        var redline = DocxDiff.Compare(left, right);

        var (mainXml, relType) = ReadMain(redline);
        Assert.Equal(TransitionalRelPrefix + "officeDocument", relType);
        Assert.Contains(TransitionalMain, mainXml);
        Assert.DoesNotContain(StrictMain, mainXml);
        Assert.Contains("<w:ins ", mainXml);
        Assert.Contains("<w:del ", mainXml);
    }

    [Fact]
    public void Compare_of_identical_strict_packages_returns_a_transitional_package()
    {
        // Word converts a strict document to transitional on open no matter what the compare finds —
        // even the no-difference result is a transitional package. The identical-package shortcut
        // must not hand back strict bytes that LibreOffice renders poorly and python-docx rejects.
        var strict = AsStrict(IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>Alpha shared body text.</w:t></w:r></w:p>"));
        var identical = new WmlDocument(strict);

        var redline = DocxDiff.Compare(strict, identical);

        var (mainXml, relType) = ReadMain(redline);
        Assert.Equal(TransitionalRelPrefix + "officeDocument", relType);
        Assert.DoesNotContain(StrictMain, mainXml);
        Assert.DoesNotContain("<w:ins ", mainXml);
    }

    // The front door reaches the same identical-package shortcut as the raw engine above, so it is
    // pinned separately: through v10 this took an engine argument, and the strict normalization lived
    // on the selector branch. With the selector gone the shortcut is the front door's own, and this
    // is what keeps it from regressing to handing back strict bytes.
    [Fact]
    public void Front_door_compare_of_identical_strict_packages_returns_a_transitional_package()
    {
        var strict = AsStrict(IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>Alpha shared body text.</w:t></w:r></w:p>"));
        var identical = new WmlDocument(strict);

        var redline = DocxCompare.Compare(strict, identical);

        var (mainXml, relType) = ReadMain(redline);
        Assert.Equal(TransitionalRelPrefix + "officeDocument", relType);
        Assert.DoesNotContain(StrictMain, mainXml);
    }

    [Fact]
    public void Compare_of_strict_left_and_transitional_right_yields_a_transitional_redline()
    {
        var left = AsStrict(IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>Alpha shared body text.</w:t></w:r></w:p>"));
        var right = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>Alpha revised body text.</w:t></w:r></w:p>");

        var redline = DocxDiff.Compare(left, right);

        var (mainXml, relType) = ReadMain(redline);
        Assert.Equal(TransitionalRelPrefix + "officeDocument", relType);
        Assert.DoesNotContain(StrictMain, mainXml);
        Assert.Contains("<w:ins ", mainXml);
    }
}
