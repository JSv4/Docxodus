using System.IO.Compression;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus.Tests.Ir;
using Docxodus.Verification;
using Xunit;

namespace Docxodus.Tests;

public class PackageManifestHyperlinkTests
{
    private const string Office = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string StrictOffice = "http://purl.oclc.org/ooxml/officeDocument/relationships";

    [Theory]
    [InlineData(Office, "other.docx", null)]
    [InlineData(Office, "../../other.docx", "Internal")]
    [InlineData(Office, "#section", null)]
    [InlineData(StrictOffice, "other.docx", null)]
    public void Internal_hyperlinks_are_references_without_package_part_targets(
        string vocabulary, string target, string? mode)
    {
        var bytes = RewriteLink(LinkDocument(), vocabulary + "/hyperlink", target, mode);
        var manifest = PackageManifestGenerator.Generate(bytes);

        Assert.True(manifest.IsValid, string.Join("\n", manifest.Findings));
        var link = Assert.Single(manifest.Relationships, r => r.Id == "link");
        Assert.Equal("Internal", link.TargetMode);
        Assert.Equal(target, link.Target);
        Assert.Null(link.ResolvedTargetUri);
        Assert.Null(link.IsTargetPresent);
    }

    [Fact]
    public void An_unedited_session_with_a_relative_hyperlink_passes_semantic_and_deliverable_preflight()
    {
        using var session = new DocxSession(LinkDocument());

        Assert.Empty(session.GetSemanticChanges().Changes);
        var report = session.VerifyDeliverable();
        Assert.True(report.AnalysisCompleted);
        Assert.Equal(DeliverableVerificationDecision.Passed, report.Decision);
        Assert.Empty(report.Findings);
        Assert.Contains(session.GetPackageManifest().Relationships, r => r.Target == "other.docx");
    }

    [Fact]
    public void Retargeting_a_relative_hyperlink_is_still_a_semantic_change()
    {
        var before = LinkDocument();
        var after = RewriteLink(before, Office + "/hyperlink", "revised.docx", null);

        var changes = SemanticDiff.Compare(before, after);

        Assert.Contains(changes.Changes, c => c.Family == SemanticChangeFamily.Relationship);
        Assert.Contains(changes.Changes, c => c.Family == SemanticChangeFamily.Hyperlink);
        Assert.DoesNotContain(changes.Changes, c => c.Family == SemanticChangeFamily.Text);
        Assert.Contains("other.docx", changes.ToCanonicalJson());
        Assert.Contains("revised.docx", changes.ToCanonicalJson());
    }

    [Theory]
    [InlineData(Office + "/image")]
    [InlineData(Office + "/oleObject")]
    [InlineData(Office + "/attachedTemplate")]
    [InlineData("urn:unrecognized/hyperlink")]
    public void Other_internal_relationships_still_require_package_parts(string type)
    {
        var manifest = PackageManifestGenerator.Generate(
            RewriteLink(LinkDocument(), type, "missing.bin", null));

        Assert.False(manifest.IsValid);
        Assert.Contains(manifest.Findings, f => f.Code == "missing_target"
            && f.Location?.RelationshipId == "link");
    }

    [Fact]
    public void A_hyperlink_with_an_invalid_mode_is_still_invalid()
    {
        var manifest = PackageManifestGenerator.Generate(
            RewriteLink(LinkDocument(), Office + "/hyperlink", "other.docx", "external"));

        Assert.False(manifest.IsValid);
        Assert.Contains(manifest.Findings, f => f.Code == "invalid_target_mode");
    }

    private static byte[] LinkDocument()
    {
        using var stream = new MemoryStream();
        stream.Write(IrTestDocuments.Create().DocumentByteArray);
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var main = document.MainDocumentPart!;
            main.AddHyperlinkRelationship(new Uri("other.docx", UriKind.Relative), false, "link");
            main.Document!.Body!.Append(new Paragraph(new Hyperlink(new Run(new Text("Other file")))
                { Id = "link" }));
        }
        return stream.ToArray();
    }

    private static byte[] RewriteLink(byte[] bytes, string type, string target, string? mode)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Update, true))
        {
            var entry = zip.GetEntry("word/_rels/document.xml.rels")!;
            XDocument xml;
            using (var input = entry.Open()) xml = XDocument.Load(input);
            var link = xml.Root!.Elements().Single(r => (string?)r.Attribute("Id") == "link");
            link.SetAttributeValue("Type", type);
            link.SetAttributeValue("Target", target);
            link.SetAttributeValue("TargetMode", mode);
            using var output = entry.Open();
            output.SetLength(0);
            xml.Save(output);
        }
        return stream.ToArray();
    }
}
