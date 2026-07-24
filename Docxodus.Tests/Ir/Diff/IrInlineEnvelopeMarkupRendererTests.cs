#nullable enable

using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Docxodus;
using Docxodus.Ir;
using Docxodus.Ir.Diff;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// Structural round trips for inline content-control/smart-tag carriers. The reader intentionally flattens
/// these wrappers for text alignment, so these tests assert the stronger package invariant: Accept retains the
/// complete right inline tree and Reject retains the complete left tree, even when visible text is unchanged.
/// </summary>
public class IrInlineEnvelopeMarkupRendererTests
{
    private static readonly XNamespace W = IrTestDocuments.W;

    private static WmlDocument Doc(string body) => IrTestDocuments.FromBodyXml(body);

    private static string Plain(string text = "controlled") =>
        $"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>";

    private static string Sdt(string tag, string text = "controlled") =>
        "<w:p><w:sdt><w:sdtPr><w:tag w:val=\"" + tag + "\"/></w:sdtPr>" +
        "<w:sdtContent><w:r><w:t>" + text + "</w:t></w:r></w:sdtContent></w:sdt></w:p>";

    private static string SmartTag(string text = "controlled") =>
        "<w:p><w:smartTag w:uri=\"urn:schemas-microsoft-com:office:smarttags\" " +
        "w:element=\"country-region\"><w:r><w:t>" + text +
        "</w:t></w:r></w:smartTag></w:p>";

    [Fact]
    public void Render_InlineSdtAdditionAndRemoval_RoundTripsExactInlineShape()
    {
        var plain = Doc(Plain());
        var wrapped = Doc(Sdt("new"));

        AssertRoundTrips(plain, wrapped);
        AssertRoundTrips(wrapped, plain);

        var revisions = DocxDiff.GetRevisions(plain, wrapped);
        Assert.Contains(revisions, revision => revision.Type == DocxDiffRevisionType.Deleted && revision.Text == "controlled");
        Assert.Contains(revisions, revision => revision.Type == DocxDiffRevisionType.Inserted && revision.Text == "controlled");
    }

    [Fact]
    public void Render_InlineSdtMetadataOnlyChange_RoundTripsOldAndNewEnvelope()
    {
        AssertRoundTrips(Doc(Sdt("old")), Doc(Sdt("new")));
    }

    [Fact]
    public void Render_SmartTagWrapperChange_RoundTripsExactInlineShape()
    {
        // w:smartTag is legacy markup that OpenXmlValidator (Office 2019) rejects even in the source fixture;
        // assert its structural accept/reject behavior here, while the SDT cases above carry schema coverage.
        AssertRoundTrips(Doc(Plain()), Doc(SmartTag()), validateSchema: false);
    }

    [Fact]
    public void Render_InlineSdtInnerTextChange_UsesWholeParagraphFallback()
    {
        var left = Doc(Sdt("stable", "old payload"));
        var right = Doc(Sdt("stable", "new payload"));
        var script = IrEditScriptBuilder.Build(IrReader.Read(left), IrReader.Read(right), new IrDiffSettings());
        var op = Assert.Single(script.Operations);
        Assert.True(op.RequiresWholeParagraphReplace);

        AssertRoundTrips(left, right);
    }

    [Fact]
    public void Consolidate_SingleInlineSdtEnvelopeChange_RoundTripsAndSerializesStructuralFlag()
    {
        var left = Doc(Sdt("old"));
        var right = Doc(Sdt("new"));
        var reviewers = new[] { new DocxDiffReviewer { Author = "Alice", Document = right } };

        var merged = DocxDiff.Consolidate(left, reviewers);
        AssertSchemaValid(merged);

        AssertInlineTreeEqual(right, RevisionProcessor.AcceptRevisions(merged));
        AssertInlineTreeEqual(left, RevisionProcessor.RejectRevisions(merged));
        Assert.Contains("\"requiresWholeParagraphReplace\"",
            DocxDiff.GetConsolidatedEditScriptJson(left, reviewers));

        var revisions = DocxDiff.GetConsolidatedRevisions(left, reviewers);
        Assert.Contains(revisions, revision =>
            revision.Type == DocxDiffRevisionType.Deleted &&
            revision.Text == "controlled" &&
            revision.Author == "Alice");
        Assert.Contains(revisions, revision =>
            revision.Type == DocxDiffRevisionType.Inserted &&
            revision.Text == "controlled" &&
            revision.Author == "Alice");
    }

    private static void AssertRoundTrips(WmlDocument left, WmlDocument right, bool validateSchema = true)
    {
        var redline = DocxDiff.Compare(left, right);
        if (validateSchema)
            AssertSchemaValid(redline);
        Assert.Contains(MainXml(redline).Descendants(W + "ins"), _ => true);
        Assert.Contains(MainXml(redline).Descendants(W + "del"), _ => true);

        var accepted = RevisionProcessor.AcceptRevisions(redline);
        var rejected = RevisionProcessor.RejectRevisions(redline);

        AssertInlineTreeEqual(right, accepted);
        AssertInlineTreeEqual(left, rejected);
        AssertNoRevisionMarkup(accepted);
        AssertNoRevisionMarkup(rejected);
    }

    private static void AssertSchemaValid(WmlDocument doc)
    {
        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(stream, false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(wdoc)
            .Select(error => $"{error.Id}@{error.Path?.XPath}: {error.Description}")
            .ToList();
        Assert.True(errors.Count == 0, string.Join("\n", errors));
    }

    private static void AssertInlineTreeEqual(WmlDocument expected, WmlDocument actual) =>
        Assert.True(XNode.DeepEquals(InlineTree(expected), InlineTree(actual)),
            $"expected: {InlineTree(expected)}\nactual: {InlineTree(actual)}");

    private static XElement InlineTree(WmlDocument doc)
    {
        var paragraph = MainXml(doc).Root!.Element(W + "body")!.Elements(W + "p").Single();
        var tree = new XElement("inline",
            paragraph.Elements().Where(element => element.Name != W + "pPr").Select(CloneWithoutNoise));
        return tree;
    }

    private static XElement CloneWithoutNoise(XElement source)
    {
        var clone = new XElement(source);
        foreach (var attribute in clone.DescendantsAndSelf().Attributes().Where(attribute =>
                     attribute.Name.LocalName.StartsWith("rsid", System.StringComparison.Ordinal) ||
                     attribute.Name.NamespaceName == "http://powertools.codeplex.com/2011" ||
                     attribute.Name.NamespaceName == "http://powertools.codeplex.com/documentbuilder/2011").ToList())
            attribute.Remove();
        return clone;
    }

    private static void AssertNoRevisionMarkup(WmlDocument doc)
    {
        var revisions = new HashSet<XName>
        {
            W + "ins", W + "del", W + "moveFrom", W + "moveTo",
            W + "moveFromRangeStart", W + "moveFromRangeEnd", W + "moveToRangeStart", W + "moveToRangeEnd",
        };
        Assert.DoesNotContain(MainXml(doc).Descendants(), element => revisions.Contains(element.Name));
    }

    private static XDocument MainXml(WmlDocument doc)
    {
        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(stream, false);
        using var reader = new StreamReader(wdoc.MainDocumentPart!.GetStream());
        return XDocument.Parse(reader.ReadToEnd());
    }
}
